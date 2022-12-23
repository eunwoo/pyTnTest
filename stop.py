from serialcom import commaster

LoadInverter = commaster.iV5Inverter(port = "COM2", baud = 9600)
LoadInverter.SetStationID(2)
LoadInverter.WriteRunStopSrc(b'0004')
LoadInverter.Stop()

TestInverter = commaster.S100Inverter(port = "COM11", baud = 19200)
TestInverter.SetStationID(2)
TestInverter.Stop()

