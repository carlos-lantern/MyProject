def is_prime(n):
    if n < 2:
        return False
    if n in (2, 3):
        return True
    if n % 2 == 0:
        return False
    for i in range(3, int(n ** 0.5) + 1, 2):
        if n % i == 0:
            return False
    return True


def find_primes(limit):
    return [n for n in range(2, limit + 1) if is_prime(n)]


if __name__ == "__main__":
    print(find_primes(100))
