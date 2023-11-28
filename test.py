import asyncio
from random import random

# async def task_coro(arg):
#     value = random()
#     print(f'>task got {value}')
#     await asyncio.sleep(value)
#     print('>task done')
#
# async def main():
#     task = task_coro(1)
#     asyncio.wait_for(task, timeout=1)  # Здесь происходит ошибка, потому что мы не используем await
#
# asyncio.run(main())



async def task_sleep():
    await asyncio.sleep(1)
    print(1)



async def main():
    coro = [task_sleep() for i in range(100)]
    return await asyncio.gather(*coro)

asyncio.run(main())