from collections import deque

def brackets(world):
    stek = deque()
    for i in world:
        if i == '(':
            stek.appendleft(i)
        elif  i == ')':
                if stek.count('(') != 0:
                    stek.popleft()
                else:
                    return False
    return not stek

print(brackets("(()())"))
# True
print(brackets(""))
# True
print(brackets("(()()))"))
# False