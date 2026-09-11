"""Admin fee and total charge calculations."""

from typing import Tuple


def calculate_toll_totals(
    toll_amount: float, admin_fee: float
) -> Tuple[float, float, float]:
    """
    Calculate positive toll amount, admin fee, and total customer amount.
    Returns: (positive_toll_amount, admin_fee, total_amount)
    """
    pos_toll = round(abs(float(toll_amount)), 2)
    fee = round(abs(float(admin_fee)), 2)
    total = round(pos_toll + fee, 2)
    return pos_toll, fee, total
