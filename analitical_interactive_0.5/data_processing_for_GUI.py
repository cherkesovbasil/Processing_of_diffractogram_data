var = 123
for X in range(4):
    globals()['matrix%d' % X] = var
