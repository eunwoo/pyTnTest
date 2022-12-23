def hex2ascii(hex):
    if hex >= 0 and hex <= 9:
        return hex + 0x30
    else:
        return hex - 10 + 0x41

def checksum(data):
    sum = 0
    for i in data:
        sum += i
    x = bytearray()
    x.append(hex2ascii(int((sum & 0xf0)/16)))
    x.append(hex2ascii((sum & 0xf)))
    return x

if __name__ == 'main':
    addr = b'7520'
    data = b'0064'
    txdata = addr + b'1' + data
    print(checksum(txdata))

def crc16(data):
    M16 = int.from_bytes(b'\xA0\x01', 'big')
    crc = int.from_bytes(b'\xff\xff', 'big')
    print(crc)
    for byte in data:
        crc = crc ^ byte
        print(byte, crc)
        for i in range(8):
            if (crc & 1) == 1:
                carrybit = 1
            else:
                carrybit = 0
            crc >>= 1
            if carrybit:
                crc ^= M16
    print(crc)
    print(int.to_bytes(crc, 2, 'little'))
    return int.to_bytes(crc, 2, 'little')
# WORD CRC16 (BYTE *buf, BYTE numbytes)	//2006.08.10 LBK sci에서 이동
# {
#     WORD crc;
#     BYTE i;
#     BYTE carrybit;
#     BYTE *ptr = buf;
#     crc = 0xffff;
#     do
#     {
#         crc = crc ^ ((WORD)*ptr++);
#         i = 8;
#         do
#         {
#             carrybit = (crc & 0x0001) ? 1 : 0;
#             crc >>= 1;
# 	        if (carrybit) crc ^= M16;
#         } while(--i);
#     } while (--numbytes);
    
#    return crc;
# }
