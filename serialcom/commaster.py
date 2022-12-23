#-*- coding:utf-8 -*-
import sys
import time
import serial.tools.list_ports
import threading
import queue
from . import checksum
from PySide6.QtCore import QObject, QTimer
import re
import copy
from queue import Queue
from serial.serialutil import SerialTimeoutException

def scan():
    port_and_desc = []
    ports = serial.tools.list_ports.comports()
    for port, desc, hwid in sorted(ports):
        # print("{}: {} [{}]".format(port, desc, hwid))
        one_port = {}
        one_port['PORT'] = port
        one_port['DESC'] = desc
        port_and_desc.append(one_port) 
    return port_and_desc

class Protocol(object):
    def __init__(self):
        self.timeout = False
        self.waiting = False
    """\
    Protocol as used by the ReaderThread. This base class provides empty
    implementations of all methods.
    """

    def connection_made(self, transport):
        """Called when reader thread is started"""

    def data_received(self, data):
        """Called with snippets received from the serial port"""

    def connection_lost(self, exc):
        """\
        Called when the serial port is closed or the reader loop terminated
        otherwise.
        """
        if isinstance(exc, Exception):
            raise exc
    
    def timeout_occurred(self):
        self.timeout = True
        """Called when timeout occurred"""

    def End(self):
        self.waiting = False
        self.received = False

class ReaderThread(threading.Thread):
    """\
    Implement a serial port read loop and dispatch to a Protocol instance (like
    the asyncio.Protocol) but do it with threads.
    Calls to close() will close the serial port but it is also possible to just
    stop() this thread and continue the serial port instance otherwise.
    """

    def __init__(self, serial_instance, protocol_factory, instance=None):
        """\
        Initialize thread.
        Note that the serial_instance' timeout is set to one second!
        Other settings are not changed.
        """
        super(ReaderThread, self).__init__()
        self.daemon = True
        self.serial = serial_instance
        self.protocol_factory = protocol_factory
        self.alive = True
        self._lock = threading.Lock()
        self._connection_made = threading.Event()
        self.protocol = None
        self.master_instance = instance

    def stop(self):
        """Stop the reader thread"""
        self.alive = False
        if hasattr(self.serial, 'cancel_read'):
            self.serial.cancel_read()
        self.join(2)

    def run(self):
        """Reader loop"""
        print('reader loop running')
        if not hasattr(self.serial, 'cancel_read'):
            self.serial.timeout = 1
        self.protocol = self.protocol_factory(self.master_instance)
        try:
            self.protocol.connection_made(self)
        except Exception as e:
            self.alive = False
            self.protocol.connection_lost(e)
            self._connection_made.set()
            return
        error = None
        self._connection_made.set()
        while self.alive and self.serial.is_open:
            try:
                # read all that is there or wait for one byte (blocking)
                # if self.master_instance.__class__.__name__ == 'TorqueMeter':
                # print(f'[{time.time()}] read(start)', self.master_instance.__class__.__name__)
                data = self.serial.read(self.serial.in_waiting or 1)
            except serial.SerialException as e:
                # probably some I/O problem such as disconnected USB serial
                # adapters -> exit
                print('serial.SerialException occurred')
                error = e
                break
            else:
                if data:
                    # make a separated try-except for called used code
                    try:
                        # if self.master_instance.__class__.__name__ == 'S100Inverter':
                        #     print(f'[{time.time()}] data', self.master_instance.__class__.__name__, data)
                        self.protocol.data_received(data)
                    except Exception as e:
                        print('exception occurred')
                        print(data)
                        print(e)
                        error = e
                        break
                else:
                    if self.protocol.waiting:
                        self.protocol.timeout_occurred()
                        if self.master_instance.__class__.__name__ == 'TorqueMeter':
                            print(f'[{time.time()}] timeout', self.master_instance.__class__.__name__)
                    
        self.alive = False
        self.protocol.connection_lost(error)
        print('protocol set to none')
        self.protocol = None

    def write(self, data):
        """Thread safe writing (uses lock)"""
        with self._lock:
            if self.master_instance.__class__.__name__ == 'TorqueMeter':
                print(f'[{time.time()}] write', self.master_instance.__class__.__name__, data)
            # print(data)
            self.serial.write(data)

    def close(self):
        """Close the serial port and exit reader thread (uses lock)"""
        # use the lock to let other threads finish writing
        with self._lock:
            # first stop reading, so that closing can be done on idle port
            self.stop()
            self.serial.close()

    def connect(self):
        """
        Wait until connection is set up and return the transport and protocol
        instances.
        """
        if self.alive:
            self._connection_made.wait()
            if not self.alive:
                raise RuntimeError('connection_lost already called')
            return (self, self.protocol)
        else:
            raise RuntimeError('already stopped')

    # - -  context manager, returns protocol

    def __enter__(self):
        """\
        Enter context handler. May raise RuntimeError in case the connection
        could not be created.
        """
        self.start()
        self._connection_made.wait()
        if not self.alive:
            raise RuntimeError('connection_lost already called')
        return self.protocol

    def __exit__(self, exc_type=0, exc_val=0, exc_tb=0):
        """Leave context: close port"""
        self.close()

# 프로토콜
class TS2700TorqueMeterProtocol(Protocol):
    def __init__(self, master_instance=None):
        super().__init__()

    # 연결 시작시 발생
    def connection_made(self, transport):
        self.transport = transport
        self.running = True
        self.rxBuf = bytearray()
        self.received = True

    # 연결 종료시 발생
    def connection_lost(self, exc):
        self.transport = None

    #데이터가 들어오면 이곳에서 처리함.
    def data_received(self, data):
        # print(data)
        if b'\n' in  data:
            for one_data in data:
                if one_data <= 127:
                    one_byte = one_data.to_bytes(1, byteorder='big')
                    self.rxBuf.extend(one_byte)
            self.rxStr = self.rxBuf.decode('ascii')
            print(self.rxStr)
            self.received = True
        else:
            for one_data in data:
                if one_data <= 127:
                    one_byte = one_data.to_bytes(1, byteorder='big')
                    self.rxBuf.extend(one_byte)

    # 데이터 보낼 때 함수
    def write(self,data):
        # print(data)
        self.timeout = False
        self.received = False
        self.rxBuf = bytearray()
        try:
            self.transport.write(data)
        except SerialTimeoutException:
            self.write(data)    # retry

    # 종료 체크
    def isDone(self):
        return self.running

class LS485Protocol(Protocol):
    class Msg:
        STATE_CREATED = 0
        STATE_RECEIVED = 1

        def __init__(self, msgID, rawdata, callback, device_name):
            self.msgID = msgID
            self.rawdata = rawdata
            self.rxMsgRaw = None
            self.rxMsg = None
            self.callback = callback
            self.state = LS485Protocol.Msg.STATE_CREATED
            self.device_name = device_name

    def __init__(self, master_instance=None):
        super().__init__()
        self.msgq = []

    # 연결 시작시 발생
    def connection_made(self, transport):
        self.transport = transport
        self.running = True
        self.rxBuf = bytearray()
        self.received = True

    # 연결 종료시 발생
    def connection_lost(self, exc):
        self.transport = None

    #데이터가 들어오면 이곳에서 처리함.
    def data_received(self, data):
        if b'\x04' in  data:
            for one_data in data:
                if one_data <= 127:
                    one_byte = one_data.to_bytes(1, byteorder='big')
                    self.rxBuf.extend(one_byte)
            if self.rxBuf.find(b'\x06') != -1:  # 0x06 = ACK
                self.rxStr = self.rxBuf[self.rxBuf.find(b'\x06'):].decode('ascii')
                print(self.rxBuf)
                self.received = True
                print('LS485Protocol received')
                msg = self.GetNextMsg()
                if msg:
                    msg.rxMsgRaw = copy.deepcopy(self.rxBuf)
                    msg.rxMsg = copy.deepcopy(self.rxStr)
                    msg.state = LS485Protocol.Msg.STATE_RECEIVED
                    self.PrintMsg(-5)
                    msg = self.GetNextMsg()
                    if msg:
                        self.write(msg.rawdata)
        else:
            for one_data in data:
                if one_data <= 127:
                    one_byte = one_data.to_bytes(1, byteorder='big')
                    self.rxBuf.extend(one_byte)

    # 데이터 보낼 때 함수
    def write(self, data):
        print('LS485Protocol::write', data)
        self.timeout = False
        self.received = False
        self.waiting = True
        self.rxBuf = bytearray()
        self.transport.write(data)

    def Write(self, msgID, id, addr, data, callback=None):
        print('LS485Protocol::Write', id)
        cmd = b'W'
        len = b'1'
        payload = id + cmd + addr + len + data
        chksum = checksum.checksum(payload)
        rawdata = b'\x05' + payload + chksum + b'\x04'
        # if self.MsgQPending() == False:
        #     self.write(rawdata)
        while self.MsgQPending():
            time.sleep(0.1)
        self.write(rawdata)
        self.msgq.append(LS485Protocol.Msg(msgID, rawdata, callback, self.transport.master_instance.__class__.__name__))
        while self.MsgQPending():
            time.sleep(0.1)
            if self.timeout:
                print(self.__class__.__name__, 'TimeoutError')
                raise TimeoutError()
        self.End()

    def Read(self, msgID, id, addr, callback=None):
        cmd = b'R'
        len = b'1'
        payload = id + cmd + addr + len
        chksum = checksum.checksum(payload)
        rawdata = b'\x05' + payload + chksum + b'\x04'
        while self.MsgQPending():
            time.sleep(0.1)
        self.write(rawdata)
        self.msgq.append(LS485Protocol.Msg(msgID, rawdata, callback, self.transport.master_instance.__class__.__name__))
        while self.MsgQPending():
            time.sleep(0.1)
            if self.timeout:
                print(self.__class__.__name__, 'TimeoutError')
                raise TimeoutError()
        self.End()

    def GetRxMsg(self, msgID):
        for msg in self.msgq:
            if msg.msgID == msgID:
                return msg.rxMsg
        return None

    def MsgQPending(self):
        if len(self.msgq) == 0:
            return False
        if self.msgq[-1].state == LS485Protocol.Msg.STATE_CREATED:
            return True
        return False

    def GetNextMsg(self):
        if len(self.msgq) == 0:
            return None
        for msg in self.msgq:
            if msg.state == LS485Protocol.Msg.STATE_CREATED:
                return msg
        return None

    def PrintMsg(self, view_range = 0):
        if view_range == 0:
            for msg in self.msgq:
                print(f'MsgID:{msg.msgID}, device:{msg.device_name}, tx:{msg.rawdata}, rxRaw:{msg.rxMsgRaw}, rx:{msg.rxMsg}')
        else:
            for msg in self.msgq[view_range:]:
                print(f'MsgID:{msg.msgID}, device:{msg.device_name}, tx:{msg.rawdata}, rxRaw:{msg.rxMsgRaw}, rx:{msg.rxMsg}')
    # 종료 체크
    def isDone(self):
        return self.running

class ModbusProtocol(Protocol):
    def __init__(self, master_instance=None):
        super().__init__()

    # 연결 시작시 발생
    def connection_made(self, transport):
        self.transport = transport
        self.running = True
        self.rxBuf = bytearray()
        self.received = True

    # 연결 종료시 발생
    def connection_lost(self, exc):
        self.transport = None

    #데이터가 들어오면 이곳에서 처리함.
    def data_received(self, data):
        if b'\x04' in  data:
            for one_data in data:
                if one_data <= 127:
                    one_byte = one_data.to_bytes(1, byteorder='big')
                    self.rxBuf.extend(one_byte)
            if self.rxBuf.find(b'\x06') != -1:
                self.rxStr = self.rxBuf[self.rxBuf.find(b'\x06'):].decode('ascii')
                print(self.rxBuf)
                self.received = True
        else:
            for one_data in data:
                if one_data <= 127:
                    one_byte = one_data.to_bytes(1, byteorder='big')
                    self.rxBuf.extend(one_byte)

    # 데이터 보낼 때 함수
    def write(self, data):
        print(data)
        self.timeout = False
        self.received = False
        self.waiting = True
        self.rxBuf = bytearray()
        self.transport.write(data)

    # 종료 체크
    def isDone(self):
        return self.running

class SerialDevice:
    Buf = {}

    def __init__(self, port, baud, protocol_factory, instance):
        self.port = port
        self.baud = baud
        self.ser = serial.Serial(port, baud, parity=serial.PARITY_NONE,
                        stopbits=serial.STOPBITS_ONE, bytesize=serial.EIGHTBITS,
                        timeout=10)
        self.thread = ReaderThread(self.ser, protocol_factory, instance)
        self.thread.__enter__()
        self.MSG_ID = 0
    
    def timeOutHandler(self):
        print('timeOutHandler')
        raise TimeoutError()

class TorqueMeter(SerialDevice):
    def __init__(self, port, baud):
        super().__init__(port, baud, TS2700TorqueMeterProtocol, self)

    def Read(self):
        self.thread.protocol.write(b'RDD\r\n')
        self.thread.protocol.timeout = False
        while self.thread.protocol.received == False and self.thread.protocol.timeout == False:
            time.sleep(0.1)
        self.thread.protocol.received = False
        if self.thread.protocol.timeout:
            print(f'[{time.time()}] timeout', self.__class__.__name__, 'TimeoutError')
            raise TimeoutError()
        else:
            return self.thread.protocol.rxStr
    
class Inverter(SerialDevice):
    PROTOCOL_LS485 = 0
    PROTOCOL_MODBUS = 1

    def __init__(self, port, baud, instance):
        super().__init__(port, baud, LS485Protocol, instance)
        self.inv_protocol = Inverter.PROTOCOL_LS485

    # def __del__(self):
    #     self.thread.__exit__()

    def SetStationID(self, id):
        self.id = f'{id:02}'.encode()
        print('SetStationID', self.__class__, self.id)

    def NewMsg(self):
        self.MSG_ID += 1
        return self.MSG_ID

    def Write(self, addr, data, msgID, callback = None):
        self.thread.protocol.Write(msgID, self.id, addr, data, callback)

    def Read(self, addr, msgID, callback = None):
        if self.__class__.__name__ == 'iV5Inverter':
            print('Read',self.__class__.__name__, 'addr', addr)
        print('self.inv_protocol=', self.inv_protocol)
        if self.inv_protocol == self.PROTOCOL_LS485:
            self.thread.protocol.Read(msgID, self.id, addr, callback)
        else:
            msg = b'\x02\x03\x11\x01\x00\x01'
            chksum = checksum.crc16(msg)
            rawdata = msg+chksum
            print(rawdata)
            self.thread.protocol.write(rawdata)
            self.thread.protocol.timeout = False
            while self.thread.protocol.received == False and self.thread.protocol.timeout == False:
                time.sleep(0.1)
            if self.thread.protocol.received:
                self.Buf[msgID] = copy.deepcopy(self.thread.protocol.rxStr)
            self.End()
            if self.thread.protocol.timeout:
                print(self.__class__.__name__, 'TimeoutError')
                raise TimeoutError()


class iV5Inverter(Inverter):
    FWD = 0
    REV = 1

    CONTROL_MODE_SPEED = b'0001'
    CONTROL_MODE_TORQUE = b'0002'

    ADDR_COMMAND_RUN = b'0006'

    # iS7 확장 공통 영역
    ADDR_STATUS_SPEED_HZ = b'0311'
    ADDR_STATUS_SPEED_RPM = b'0312'

    # 파라미터
    ADDR_PARA_RUNSTOP_SRC = b'7401'
    ADDR_PARA_TORQUE_REF = b'7520'
    ADDR_PARA_SPEED0 = b'740C'
    ADDR_PARA_CONTROL_MODE = b'7501'


    def __init__(self, port, baud):
        super().__init__(port, baud, self)

    def WriteControlMode(self, data):
        msgID = self.NewMsg()
        self.Write(iV5Inverter.ADDR_PARA_CONTROL_MODE, data, msgID)

    def WriteTorqueRef(self, fvalue):
        msgID = self.NewMsg()
        value = int(fvalue*10)
        data = ConvertInt2Hex(value).encode()
        self.Write(iV5Inverter.ADDR_PARA_TORQUE_REF, data, msgID)

    def WriteSpeed0(self, fvalue):
        msgID = self.NewMsg()
        value = int(fvalue*10)
        data = ConvertInt2Hex(value).encode()
        self.Write(iV5Inverter.ADDR_PARA_SPEED0, data, msgID)

    def WriteRunStopSrc(self, data):
        msgID = self.NewMsg()
        self.Write(iV5Inverter.ADDR_PARA_RUNSTOP_SRC, data, msgID)
    
    def ParseValue(self, str):
        return int(str[4:8], 16)

    def GetReadBuf(self, msgID):
        return self.thread.protocol.GetRxMsg(msgID)

    def ReadSpeedHz(self):
        msgID = self.NewMsg()
        self.Read(iV5Inverter.ADDR_STATUS_SPEED_HZ, msgID)
        return ConvertUINT16toINT16(self.ParseValue(self.GetReadBuf(msgID)))*0.01
        
    def ReadSpeedRpm(self):
        msgID = self.NewMsg()
        print('iV5 - ReadSpeedRpm')
        self.Read(iV5Inverter.ADDR_STATUS_SPEED_RPM, msgID)
        return ConvertUINT16toINT16(self.ParseValue(self.GetReadBuf(msgID)))

    def Run(self, direction):
        msgID = self.NewMsg()
        if direction == iV5Inverter.FWD:
            data = b'0002'
        else:
            data = b'0004'
        self.Write(iV5Inverter.ADDR_COMMAND_RUN, data, msgID)

    def Stop(self):
        msgID = self.NewMsg()
        data = b'0001'
        self.Write(iV5Inverter.ADDR_COMMAND_RUN, data, msgID)


class S100Inverter(Inverter):
    YES = 1
    NO = 0

    ADDR_COMMAND_RUN = b'0006'
    ADDR_PARA_RUNSTOP_SRC = b'1106'

    # 확장 공통 영역 (0x300~)
    ADDR_INV_CAP = b'0301'
    ADDR_INPUT_VOLT = b'0302'
    ADDR_STATUS_CURRENT = b'0310'
    ADDR_STATUS_SPEED_HZ = b'0311'
    ADDR_STATUS_SPEED_RPM = b'0312'
    ADDR_COMMON_OUTPUT_VOLT = b'0314'
    ADDR_COMMON_HZRPM_UNIT = b'031D'
    ADDR_COMMON_COMMAND_SPEED = b'0380'
    # 파라미터
    # - DRV
    ADDR_PARA_CMD_FREQ = b'1101'
    ADDR_PARA_CMD_TORQUE = b'1102'
    ADDR_PARA_DRV_FREQ_REF_SRC = b'1107'
    ADDR_PARA_TORQUE_CONTROL = b'110A'
    ADDR_PARA_MOTOR_CAP = b'110E'
    # - BAS
    ADDR_PARA_SLIP = b'120C'
    ADDR_PARA_RATED_CURR = b'120D'
    ADDR_PARA_IF = b'120E'
    ADDR_PARA_RS = b'1215'
    ADDR_PARA_LSIGMA = b'1216'
    ADDR_PARA_LS = b'1217'
    ADDR_PARA_TR = b'1218'

    MotorCapID2kW = {0:0.2, 1:0.4, 2:0.75, 3:1.1, 4:1.5, 5:2.2, 6:3.0, 7:3.7, 8:4.0, 9:5.5}

    def __init__(self, port, baud):
        super().__init__(port, baud, self)

    def SetProtocol(self, protocol):
        print('SetProtocol', protocol)
        if protocol != self.inv_protocol:
            self.inv_protocol = protocol
            self.thread.__exit__()
            self.ser = serial.Serial(self.port, self.baud, parity=serial.PARITY_NONE,
                            stopbits=serial.STOPBITS_ONE, bytesize=serial.EIGHTBITS,
                            timeout=10)
            if self.inv_protocol == Inverter.PROTOCOL_LS485:
                self.thread = ReaderThread(self.ser, LS485Protocol, self)
            else:
                self.thread = ReaderThread(self.ser, ModbusProtocol, self)
            self.thread.__enter__()


    def WriteTorqueControlMode(self, sel):
        msgID = self.NewMsg()
        if sel == S100Inverter.YES:
            data = ConvertInt2Hex(1).encode()
            self.Write(S100Inverter.ADDR_PARA_TORQUE_CONTROL, data, msgID)
        else:
            data = ConvertInt2Hex(0).encode()
            self.Write(S100Inverter.ADDR_PARA_TORQUE_CONTROL, data, msgID)

    def WriteSlip(self, fvalue):
        msgID = self.NewMsg()
        value = int(fvalue)
        data = ConvertInt2Hex(value).encode()
        self.Write(S100Inverter.ADDR_PARA_SLIP, data, msgID)

    def WriteRatedCurr(self, fvalue):
        msgID = self.NewMsg()
        value = int(fvalue*10)
        print(value)
        data = ConvertInt2Hex(value).encode()
        self.Write(S100Inverter.ADDR_PARA_RATED_CURR, data, msgID)

    def WriteSpeedRef(self, fvalue):
        print('WriteSpeedRef')
        msgID = self.NewMsg()
        value = int(fvalue*100)
        data = ConvertInt2Hex(value).encode()
        self.Write(S100Inverter.ADDR_COMMON_COMMAND_SPEED, data, msgID)

    def WriteSpeedRefKpd(self, fvalue):
        print('WriteSpeedRef')
        msgID = self.NewMsg()
        value = int(fvalue*100)
        data = ConvertInt2Hex(value).encode()
        self.Write(S100Inverter.ADDR_PARA_CMD_FREQ, data, msgID)

    def WriteTorqueRef(self, fvalue):
        print('WriteTorqueRef')
        msgID = self.NewMsg()
        value = int(fvalue*10)
        data = ConvertInt2Hex(value).encode()
        self.Write(S100Inverter.ADDR_PARA_CMD_TORQUE, data, msgID)

    def WriteRunStopSrc(self, data):
        msgID = self.NewMsg()
        self.Write(S100Inverter.ADDR_PARA_RUNSTOP_SRC, data, msgID)

    def WriteFreqRefSrc(self, data):
        msgID = self.NewMsg()
        self.Write(S100Inverter.ADDR_PARA_DRV_FREQ_REF_SRC, data, msgID)

    def GetReadBuf(self, msgID):
        return self.thread.protocol.GetRxMsg(msgID)

    def ReadSpeedRefKpd(self):
        print('ReadSpeedRef')
        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_PARA_CMD_FREQ, msgID)
        value = self.ParseValue(self.GetReadBuf(msgID))*0.01
        print(value)
        return value

    def ReadRatedCurr(self):
        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_PARA_RATED_CURR, msgID)
        return self.ParseValue(self.GetReadBuf(msgID))*0.1

    def ReadHzRpmUnit(self):
        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_COMMON_HZRPM_UNIT, msgID)
        return self.ParseValue(self.GetReadBuf(msgID))
    
    def ReadSpeedHz(self):
        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_STATUS_SPEED_HZ, msgID)
        return self.ParseValue(self.GetReadBuf(msgID))

    def ReadSpeedRpm(self):
        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_STATUS_SPEED_RPM, msgID)
        return self.ParseValue(self.GetReadBuf(msgID))

    def ReadCurrent(self):
        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_STATUS_CURRENT, msgID)
        return self.ParseValue(self.GetReadBuf(msgID))*0.1

    def ReadOutputVolt(self):
        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_COMMON_OUTPUT_VOLT, msgID)
        return self.ParseValue(self.GetReadBuf(msgID))

    def ReadIf(self):
        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_PARA_IF, msgID)
        return self.ParseValue(self.GetReadBuf(msgID))*0.1

    def ReadRs(self):
        try:
            motorCap = self.MotorCap
        except AttributeError:
            self.ReadMotorCap()
        try:
            invCap = self.InvCap
        except AttributeError:
            self.ReadInvCapacity()
        try:
            invVolt = self.InvVoltClass
        except AttributeError:
            self.ReadInputVolt()

        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_PARA_RS, msgID)
        return self.ParseValue(self.GetReadBuf(msgID))*self.ParaScale('Rs')

    def ReadLsigma(self):
        try:
            motorCap = self.MotorCap
        except AttributeError:
            self.ReadMotorCap()
        try:
            invCap = self.InvCap
        except AttributeError:
            self.ReadInvCapacity()
        try:
            invVolt = self.InvVoltClass
        except AttributeError:
            self.ReadInputVolt()

        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_PARA_LSIGMA, msgID)
        return self.ParseValue(self.GetReadBuf(msgID))*self.ParaScale('Lsigma')

    def ReadLs(self):
        try:
            motorCap = self.MotorCap
        except AttributeError:
            self.ReadMotorCap()
        try:
            invCap = self.InvCap
        except AttributeError:
            self.ReadInvCapacity()
        try:
            invVolt = self.InvVoltClass
        except AttributeError:
            self.ReadInputVolt()

        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_PARA_LS, msgID)
        return self.ParseValue(self.GetReadBuf(msgID))*self.ParaScale('Ls')

    def ReadTr(self):
        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_PARA_TR, msgID)
        return self.ParseValue(self.GetReadBuf(msgID))

    def ParaScale(self, para_name):
        try:
            motorCap = self.MotorCap
        except AttributeError:
            self.ReadMotorCap()
        try:
            invCap = self.InvCap
        except AttributeError:
            self.ReadInvCapacity()
        try:
            invVolt = self.InvVoltClass
        except AttributeError:
            self.ReadInputVolt()

        if para_name == 'Rs':
            if self.InvVoltClass == 4:
                if self.MotorCap <= 4:
                    return 0.01
                elif self.MotorCap <= 12:
                    return 0.001
                else:
                    return 0.1
            else:
                if self.MotorCap <= 1:
                    return 0.01
                elif self.MotorCap <= 9:
                    return 0.001
                else:
                    return 0.1
        elif para_name == 'Lsigma':
            if self.InvVoltClass == 4:
                if self.MotorCap <= 4:
                    return 0.1
                elif self.MotorCap <= 12:
                    return 0.01
                else:
                    return 0.001
            else:
                if self.MotorCap <= 1:
                    return 0.1
                elif self.MotorCap <= 9:
                    return 0.01
                else:
                    return 0.001
        elif para_name == 'Ls':
            if self.InvVoltClass == 4:
                if self.MotorCap <= 4:
                    return 1
                elif self.MotorCap <= 12:
                    return 0.1
                else:
                    return 0.01
            else:
                if self.MotorCap <= 1:
                    return 1
                elif self.MotorCap <= 9:
                    return 0.1
                else:
                    return 0.01
        else:
            return 1
    def ParaUnit(self, para_name):
        try:
            motorCap = self.MotorCap
        except AttributeError:
            self.ReadMotorCap()
        try:
            invCap = self.InvCap
        except AttributeError:
            self.ReadInvCapacity()
        try:
            invVolt = self.InvVoltClass
        except AttributeError:
            self.ReadInputVolt()

        if para_name == 'Rs':
            if self.InvVoltClass == 4:
                if self.MotorCap <= 12:
                    return 'ohm'
                else:
                    return 'mohm'
            else:
                if self.MotorCap <= 9:
                    return 'ohm'
                else:
                    return 'mohm'
        elif para_name == 'Lsigma':
            return 'mH'
        elif para_name == 'Ls':
            return 'mH'
        else:
            return ''
    def ReadInvCapacity(self):
        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_INV_CAP, msgID)
        # print(self.ParseInvCap(f'{self.ParseValue(self.thread.protocol.rxStr):04x}'))
        return self.ParseInvCap(self.ParseValue(self.GetReadBuf(msgID)))
    def ReadInputVolt(self):
        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_INPUT_VOLT, msgID)
        invVolt = (self.ParseValue(self.GetReadBuf(msgID)) & 0x0F00)
        if invVolt == 0x200:
            self.InvVoltClass = 2
        elif invVolt == 0x400:
            self.InvVoltClass = 4
        else:
            self.InvVoltClass = 1
    def ReadMotorCap(self):
        msgID = self.NewMsg()
        self.Read(S100Inverter.ADDR_PARA_MOTOR_CAP, msgID)
        print('motor cap=',self.ParseValue(self.GetReadBuf(msgID)))
        self.MotorCap = self.ParseValue(self.GetReadBuf(msgID))
        
    def ParseInvCap(self, encodedInvCap):
        # Megawatt encoding은 처리안함
        if encodedInvCap >= 0x4000:
            decoded = (encodedInvCap - 0x4000)
            frac = decoded & 0xF
            integer = (decoded & 0x3FF0) / 16
            self.InvCap = integer * 1000 + frac*100
        else:
            self.InvCap = encodedInvCap / 16
        return self.InvCap

    def Run(self):
        msgID = self.NewMsg()
        data = b'0002'
        self.Write(S100Inverter.ADDR_COMMAND_RUN, data, msgID)

    def Stop(self):
        msgID = self.NewMsg()
        data = b'0001'
        self.Write(S100Inverter.ADDR_COMMAND_RUN, data, msgID)

    def ParseValue(self, str):
        return int(str[4:8], 16)


def ConvertInt2Hex(value):
    if value >= 0:
        hex_string = f'{value:04X}'
    else:
        hex_string = f'{65536 + value:04X}'
    return hex_string

def ConvertUINT16toINT16(value):
    if value >= 32767:
        return value - 65536
    else:
        return value