def mystery (n):
    print(n)
    if n <= 1:
        return 1
    else:
        return ((n%10)*mystery(n/10))

print(mystery(4235))
