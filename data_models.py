"""
Data models for financing costs parser
"""

from typing import List, Optional, Tuple
from dataclasses import dataclass, field


@dataclass
class ConstructionInterestFees:
    """Data structure for Construction Interest & Fees section"""
    construction_loan_interest: float = 0.0
    origination_fee: float = 0.0
    credit_enhancement_fee: float = 0.0
    bond_premium: float = 0.0
    cost_of_issuance: float = 0.0
    title_recording: float = 0.0
    taxes: float = 0.0
    insurance: float = 0.0
    other_amount: float = 0.0
    other_description: str = ""
    total: float = 0.0
    
    def get_line_items_sum(self) -> float:
        """Calculate sum of all line items"""
        return (self.construction_loan_interest + self.origination_fee + 
                self.credit_enhancement_fee + self.bond_premium + 
                self.cost_of_issuance + self.title_recording + 
                self.taxes + self.insurance + self.other_amount)
    
    def is_valid(self) -> Tuple[bool, str]:
        """Validate that total equals sum of line items"""
        line_sum = self.get_line_items_sum()
        if abs(self.total - line_sum) > 0.01:  # Allow small floating point differences
            return False, f"Total ({self.total:,.2f}) does not match sum of line items ({line_sum:,.2f})"
        return True, ""


@dataclass
class PermanentFinancing:
    """Data structure for Permanent Financing section"""
    loan_origination_fee: float = 0.0
    credit_enhancement_fee: float = 0.0
    title_recording: float = 0.0
    taxes: float = 0.0
    insurance: float = 0.0
    other_amount: float = 0.0
    other_description: str = ""
    total: float = 0.0
    
    def get_line_items_sum(self) -> float:
        """Calculate sum of all line items"""
        return (self.loan_origination_fee + self.credit_enhancement_fee + 
                self.title_recording + self.taxes + self.insurance + 
                self.other_amount)
    
    def is_valid(self) -> Tuple[bool, str]:
        """Validate that total equals sum of line items"""
        line_sum = self.get_line_items_sum()
        if abs(self.total - line_sum) > 0.01:  # Allow small floating point differences
            return False, f"Total ({self.total:,.2f}) does not match sum of line items ({line_sum:,.2f})"
        return True, ""


@dataclass
class ApplicationData:
    """Complete data structure for an application"""
    application_name: str = ""
    file_path: str = ""
    construction_interest_fees: ConstructionInterestFees = field(default_factory=ConstructionInterestFees)
    permanent_financing: PermanentFinancing = field(default_factory=PermanentFinancing)
    total_units: Optional[int] = None
    total_square_feet: Optional[float] = None
    new_construction_total: Optional[float] = None
    validation_errors: List[str] = field(default_factory=list)
    validation_warnings: List[str] = field(default_factory=list)
    
    def get_combined_financing_costs(self) -> float:
        """Calculate combined financing costs"""
        return self.construction_interest_fees.total + self.permanent_financing.total
    
    def get_financing_costs_per_unit(self) -> Optional[float]:
        """Calculate financing costs per unit"""
        if self.total_units and self.total_units > 0:
            return self.get_combined_financing_costs() / self.total_units
        return None
    
    def get_financing_costs_per_sf(self) -> Optional[float]:
        """Calculate financing costs per square foot"""
        if self.total_square_feet and self.total_square_feet > 0:
            return self.get_combined_financing_costs() / self.total_square_feet
        return None
    
    def get_financing_costs_pct_hard_costs(self) -> Optional[float]:
        """Calculate financing costs as percentage of hard costs"""
        if self.new_construction_total and self.new_construction_total > 0:
            return (self.get_combined_financing_costs() / self.new_construction_total) * 100
        return None
    
    def validate(self):
        """Run all validation checks"""
        self.validation_errors = []
        self.validation_warnings = []
        
        # Validate construction interest & fees
        is_valid, error_msg = self.construction_interest_fees.is_valid()
        if not is_valid:
            self.validation_errors.append(f"Construction Interest & Fees: {error_msg}")
        
        # Validate permanent financing
        is_valid, error_msg = self.permanent_financing.is_valid()
        if not is_valid:
            self.validation_errors.append(f"Permanent Financing: {error_msg}")
        
        # Check for missing data
        if self.total_units is None:
            self.validation_warnings.append("Total units not found")
        if self.total_square_feet is None:
            self.validation_warnings.append("Total square footage not found")
        if self.new_construction_total is None:
            self.validation_warnings.append("New Construction total not found (using fallback)")

