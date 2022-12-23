# https://tinkering.xyz/async-serial/
import asyncio
import serial_asyncio


async def main(loop):
    reader, writer = await serial_asyncio.open_serial_connection(url='COM6', baudrate=9600)
    print('Reader created')
    # _, writer = await serial_asyncio.open_serial_connection(url='COM6', baudrate=9600)
    print('Writer created')
    messages = [b'RDD\n', b'RDD\n', b'RDD\n', b'RDD\n']
    sent = send(writer, messages)
    received = recv(reader)
    await asyncio.wait([asyncio.create_task(sent), asyncio.create_task(received)])


async def send(w, msgs):
    for msg in msgs:
        w.write(msg)
        print(f'sent: {msg.decode().rstrip()}')
        await asyncio.sleep(0.5)
    w.write(b'DONE\n')
    print('Done sending')


async def recv(r):
    while True:
        msg = await r.readuntil(b'\n')
        if msg.rstrip() == b'DONE':
            print('Done receiving')
            break
        print(f'received: {msg.rstrip().decode()}')


loop = asyncio.new_event_loop()
asyncio.set_event_loop(loop)
loop.run_until_complete(main(loop))
loop.close()