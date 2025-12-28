"""
opeb_rollforward/engine.py - Production GASB 75 Roll-Forward Engine

GASB 75 Compliance:
- ¶96: TOL Reconciliation disclosures  
- ¶43(a): Experience gains/losses
- ¶43(b): Assumption change effects
- ¶44: Interest cost calculation

Author: Actuarial Pipeline Project
"""

import numpy as np
from datetime import date, datetime
from typing import Dict, Optional, Union
from dataclasses import dataclass, field
import json
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


@dataclass
class PriorValuation:
    """Prior valuation results for roll-forward."""
    valuation_date: date
    total_opeb_liability: float
    tol_actives: float
    tol_retirees: float
    service_cost: float
    discount_rate_boy: float
    discount_rate_eoy: float
    avg_remaining_service_life: float = 12.0
    trend_rates: Dict[int, float] = field(default_factory=dict)
    sensitivity_dr_plus1: Optional[float] = None
    sensitivity_dr_minus1: Optional[float] = None
    client_name: str = ""
    
    @classmethod
    def from_json(cls, filepath: Union[str, Path]) -> 'PriorValuation':
        with open(filepath, 'r') as f:
            data = json.load(f)
        val_date = data.get('valuation_date')
        if isinstance(val_date, str):
            val_date = datetime.strptime(val_date, '%Y-%m-%d').date()
        trend_rates = {int(k): float(v) for k, v in data.get('trend_rates', {}).items()}
        return cls(
            valuation_date=val_date,
            total_opeb_liability=float(data.get('total_opeb_liability', 0)),
            tol_actives=float(data.get('tol_actives', 0)),
            tol_retirees=float(data.get('tol_retirees', 0)),
            service_cost=float(data.get('service_cost', 0)),
            discount_rate_boy=float(data.get('discount_rate_boy', 0.04)),
            discount_rate_eoy=float(data.get('discount_rate_eoy', 0.04)),
            avg_remaining_service_life=float(data.get('avg_remaining_service_life', 12)),
            trend_rates=trend_rates,
            sensitivity_dr_plus1=data.get('sensitivity_dr_plus1'),
            sensitivity_dr_minus1=data.get('sensitivity_dr_minus1'),
            client_name=data.get('client_name', ''),
        )
    
    @property
    def duration_estimate(self) -> float:
        if self.avg_remaining_service_life > 0 and self.total_opeb_liability > 0:
            active_pct = self.tol_actives / self.total_opeb_liability
            return active_pct * (self.avg_remaining_service_life + 10) + (1 - active_pct) * 10
        return 10.0


@dataclass
class RollForwardResults:
    """Roll-forward calculation results."""
    boy_date: date
    eoy_date: date
    boy_tol: float
    service_cost: float
    interest_cost: float
    benefit_payments: float
    expected_eoy_tol: float
    actual_eoy_tol: Optional[float] = None
    experience_gain_loss: float = 0.0
    assumption_change_effect: float = 0.0
    discount_rate_change_effect: float = 0.0
    
    def to_dict(self) -> Dict:
        return {
            'boy_tol': self.boy_tol, 'service_cost': self.service_cost,
            'interest_cost': self.interest_cost, 'benefit_payments': self.benefit_payments,
            'expected_eoy_tol': self.expected_eoy_tol, 'actual_eoy_tol': self.actual_eoy_tol,
            'experience_gain_loss': self.experience_gain_loss,
            'assumption_change': self.assumption_change_effect,
        }
    
    def get_reconciliation_table(self) -> Dict:
        return {
            'Beginning TOL': self.boy_tol, 'Service Cost': self.service_cost,
            'Interest Cost': self.interest_cost, 'Benefit Payments': -self.benefit_payments,
            'Experience (Gain)/Loss': self.experience_gain_loss,
            'Assumption Changes': self.assumption_change_effect,
            'Ending TOL': self.actual_eoy_tol or self.expected_eoy_tol,
        }


class RollForwardEngine:
    """Production GASB 75 Roll-Forward Engine."""
    
    def __init__(self, prior: PriorValuation, current_date: date,
                 benefit_payments: float = 0.0, new_discount_rate: Optional[float] = None,
                 actual_eoy_tol: Optional[float] = None, duration: Optional[float] = None):
        self.prior = prior
        self.current_date = current_date
        self.benefit_payments = benefit_payments
        self.new_discount_rate = new_discount_rate
        self.actual_eoy_tol = actual_eoy_tol
        self.duration = duration or prior.duration_estimate
    
    def calculate_interest_cost(self) -> float:
        """Interest = (BOY_TOL + SC/2 - BP/2) × BOY_Rate"""
        avg_balance = self.prior.total_opeb_liability + (self.prior.service_cost / 2) - (self.benefit_payments / 2)
        return avg_balance * self.prior.discount_rate_boy
    
    def calculate_expected_eoy_tol(self) -> float:
        return self.prior.total_opeb_liability + self.prior.service_cost + self.calculate_interest_cost() - self.benefit_payments
    
    def calculate_discount_rate_change_effect(self) -> float:
        """ΔL ≈ -Duration × L × Δr"""
        if self.new_discount_rate is None:
            return 0.0
        delta_rate = self.new_discount_rate - self.prior.discount_rate_eoy
        if abs(delta_rate) < 0.0001:
            return 0.0
        return -self.duration * self.prior.total_opeb_liability * delta_rate
    
    def run(self) -> RollForwardResults:
        interest_cost = self.calculate_interest_cost()
        expected_eoy = self.calculate_expected_eoy_tol()
        dr_effect = self.calculate_discount_rate_change_effect()
        
        experience_gl = 0.0
        if self.actual_eoy_tol is not None:
            expected_adjusted = expected_eoy + dr_effect
            experience_gl = self.actual_eoy_tol - expected_adjusted
        
        return RollForwardResults(
            boy_date=self.prior.valuation_date, eoy_date=self.current_date,
            boy_tol=self.prior.total_opeb_liability, service_cost=self.prior.service_cost,
            interest_cost=interest_cost, benefit_payments=self.benefit_payments,
            expected_eoy_tol=expected_eoy, actual_eoy_tol=self.actual_eoy_tol,
            experience_gain_loss=experience_gl, assumption_change_effect=dr_effect,
            discount_rate_change_effect=dr_effect,
        )


def create_engine(prior: PriorValuation, current_date: date, **kwargs) -> RollForwardEngine:
    return RollForwardEngine(prior, current_date, **kwargs)

def load_prior(filepath: Union[str, Path]) -> PriorValuation:
    return PriorValuation.from_json(filepath)
