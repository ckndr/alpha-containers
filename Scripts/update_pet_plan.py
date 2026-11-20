"""
update_pet_plan.py
Thin backward-compatible wrapper delegating to update_plans.main().
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import update_plans


def main():
    update_plans.main()


if __name__ == '__main__':
    main()
