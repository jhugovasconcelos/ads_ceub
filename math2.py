import math
def bhaskara(a, b, c):
    Delta = (b ** 2) - 4 * a * c

    x1 = -b + math.sqrt(Delta) / -a
    x2 = -b - math.sqrt(Delta) / -a
    return x1, x2
