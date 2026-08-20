from __future__ import annotations

import hashlib
import sys


def main() -> int:
    if len(sys.argv) != 2:
        print('Usage: python scripts/hash_password.py "your-password"')
        return 1

    print(hashlib.sha256(sys.argv[1].encode("utf-8")).hexdigest())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
