def debugPrint(*args, level=0, sep=' ', end='\n', file=None):
    if level == 0:
        print(args, sep=sep, end=end, file=file)