# ==========================================
# 1. CARBON STANDARD DATA (True VPDB Values)
# ==========================================
# Values extracted from the provided JavaScript
CARBON_STANDARDS = {
    # Group 1 Standards
    "LSVEC": -46.6,
    "NBS18": -5.01,
    "IAEA-610": -9.109,
    "IAEA-611": -30.795,
    "IAEA-612": -36.722,
    "USGS44": -42.21,
    
    # Group 2 Standards (Anchors)
    "NBS19": 1.95,
    "IAEA-603": 2.460
}

# ==========================================
# 2. CALCULATOR CLASS
# ==========================================

class CarbonIsotopeCalculator:

    @staticmethod
    def get_standard_value(std_name):
        """
        Retrieves the true VPDB value for a given standard name.
        """
        try:
            return CARBON_STANDARDS[std_name]
        except KeyError:
            raise ValueError(f"Standard '{std_name}' not found in database. Available: {list(CARBON_STANDARDS.keys())}")

    @staticmethod
    def calculate_slope(y_true_1, y_true_2, x_meas_1, x_meas_2):
        """
        Calculates Slope (m) for the normalization equation y = mx + b.
        
        y = True Value
        x = Measured Value
        """
        numerator = y_true_1 - y_true_2
        denominator = x_meas_1 - x_meas_2
        
        if denominator == 0:
            return 0 # Avoid division by zero
            
        return numerator / denominator

    @staticmethod
    def calculate_intercept(y_true, x_meas, m):
        """
        Calculates Intercept (b) for the normalization equation y = mx + b.
        """
        return y_true - (m * x_meas)

    @classmethod
    def process_sample(cls, 
                       meas_std_1, name_std_1,
                       meas_std_2, name_std_2,
                       meas_sample):
        """
        Main function to call from Excel/Pandas.
        
        Parameters:
        -----------
        meas_std_1 : float -> Measured value of the first standard (e.g., NBS19)
        name_std_1 : str   -> Name of the first standard (must match dictionary keys)
        meas_std_2 : float -> Measured value of the second standard (e.g., LSVEC)
        name_std_2 : str   -> Name of the second standard
        meas_sample: float -> Measured value of the unknown sample
        
        Returns:
        --------
        Dictionary containing corrected VPDB value, slope, and intercept.
        """
        
        # 1. Get True Values
        true_std_1 = cls.get_standard_value(name_std_1)
        true_std_2 = cls.get_standard_value(name_std_2)

        # 2. Calculate Slope (m)
        m = cls.calculate_slope(true_std_1, true_std_2, meas_std_1, meas_std_2)

        # 3. Calculate Intercept (b) using Standard 1
        b = cls.calculate_intercept(true_std_1, meas_std_1, m)

        # 4. Correct the Sample
        # Equation: y = mx + b
        sample_corrected = (m * meas_sample) + b

        return {
            "d13C_VPDB": round(sample_corrected, 2),
            "slope_m": round(m, 5),
            "intercept_b": round(b, 5),
            "std1_used": name_std_1,
            "std2_used": name_std_2
        }
    

"""
Sample Usage

import carbon_isotope_calculator as carb

# ==========================================
# EXAMPLE USAGE
# ==========================================

# 1. Define your inputs (from Excel row)
# Note: Typically one high value (like NBS19) and one low value (like LSVEC) are used.
inputs = {
    "meas_std_1": 1.85,       # Measured NBS19 (True is 1.95)
    "name_std_1": "NBS19",
    
    "meas_std_2": -46.50,     # Measured LSVEC (True is -46.6)
    "name_std_2": "LSVEC",
    
    "meas_sample": -5.20      # Measured Unknown Sample
}

# 2. Run the calculation
try:
    result = carb.CarbonIsotopeCalculator.process_sample(
        meas_std_1=inputs["meas_std_1"],
        name_std_1=inputs["name_std_1"],
        meas_std_2=inputs["meas_std_2"],
        name_std_2=inputs["name_std_2"],
        meas_sample=inputs["meas_sample"]
    )

    print("--- Carbon Calculation Results ---")
    print(f"Sample d13C (VPDB): {result['d13C_VPDB']} permil")
    print(f"Slope (m):          {result['slope_m']}")
    print(f"Intercept (b):      {result['intercept_b']}")

except Exception as e:
    print(f"Error: {e}")
"""