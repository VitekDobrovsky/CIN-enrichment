from time import sleep

def countDown(preMsg: str, postMsg: str, amount: int):
    for i in range(amount, 0, -1):
        print(preMsg + str(i) + postMsg, end="\r")
        sleep(1)
