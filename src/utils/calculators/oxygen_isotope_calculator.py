import math

# ==========================================
# 1. MINERAL DATA REPOSITORY
# ==========================================
MINERAL_DATA = {
    "temp_non_25": {
        "Ankerite": {
            "Rosenbaum and Shepared (1986)": {"w": 0.668, "x": 4.15, "eqn": "TSqr"}
        },
        "Aragonite": {
            "Kim et al. (2007)": {"w": 3.3916, "x": -0.8276, "eqn": "T"}
        },
        "Calcite": {
            "Kim et al. (2007)": {"w": 3.5902, "x": -1.7881, "eqn": "T"},
            "Sharma and Clayton (1965) + Kim et al. (2007)": {"w": 3.48, "x": -1.47, "eqn": "T"}
        },
        "Cerussite": {
            "Gilg et al. (2003)": {"w": 0.479, "x": 5.13, "eqn": "TSqr"}
        },
        "Dolomite": {
            "Rosenbaum and Shepared (1986)": {"w": 0.665, "x": 4.23, "eqn": "TSqr"}
        },
        "Magnesite": {
            "Das Sharma et al. (2002)": {"w": 0.6845, "x": 4.22, "eqn": "TSqr"}
        },
        "Siderite": {
            "Rosenbaum and Sheppard (1986)": {"w": 0.684, "x": 3.85, "eqn": "TSqr"},
            "Carothers et al. (1988) - Natural": {"w": 3.2758, "x": 0.5929, "eqn": "T"}
        },
        "Smithsonite": {
            "Gillg et al. (2003)": {"w": 0.669, "x": 3.96, "eqn": "TSqr"}
        }
    },
    "temp_25": {
        "Ankerite": {
            "Rosenbaum and Shepared (1986)": {"alpha": 1.01177}
        },
        "Aragonite": {
            "Kim et al. (2007)": {"alpha": 1.01063}
        },
        "Calcite": {
            "Sharma and Clayton (1965) corrected": {"alpha": 1.01025}
        },
        "Cerussite": {
            "Gilg et al. (2003)": {"alpha": 1.01061}
        },
        "Dolomite": {
            "Sharma and Clayton (1965) corrected": {"alpha": 1.01109}
        },
        "Otavite": {
            "Sharma and Clayton (1965) corrected": {"alpha": 1.01146}
        },
        "Protodolomite": {
            "Land (1980)": {"alpha": 1.01265}
        },
        "Rhodochrosite": {
            "Sharma and Clayton (1965) corrected": {"alpha": 1.01012}
        },
        "Siderite": {
            "Carothers et al. (1988)": {"alpha": 1.01165}
        },
        "Smithsonite": {
            "Gilg et al. (2003)": {"alpha": 1.01149}
        },
        "Strontianite": {
            "Sharma and Clayton (1965) corrected": {"alpha": 1.01048}
        },
        "Witherite": {
            "Kim and O'Neil (1997)": {"alpha": 1.01063}
        }
    }
}

# Standard Reference Values (Constants)
STD_VALUES = {
    "NBS18": {"VPDB": -23.01, "VSMOW": 7.2},
    "NBS19": {"VPDB": -2.20, "VSMOW": 28.65},
    "IAEA-603": {"VPDB": -2.37, "VSMOW": 28.48}
}

# ==========================================
# 2. CALCULATOR CLASS
# ==========================================

class OxygenIsotopeCalculator:
    
    @staticmethod
    def get_calcite_acid_fractionation(temp_c):
        """
        Calculates the AFF for Calcite based on Kim et al. (2015).
        Equation: 1000 ln(alpha) = 3.48(10^3/T) - 1.47
        """
        temp_k = temp_c + 273.15
        val = (3.48 * (1000 / temp_k)) - 1.47
        return math.exp(val / 1000)

    @staticmethod
    def get_mineral_alpha(mineral, citation, temp_c, user_defined_alpha=None):
        """
        Determines the Acid Fractionation Factor (alpha) based on mineralogy.
        """
        if user_defined_alpha is not None:
            return float(user_defined_alpha)

        temp_k = temp_c + 273.15
        
        # Check if strictly 25C logic should be applied
        # (Using a small epsilon for float comparison, or exactly 25)
        if abs(temp_c - 25.0) < 0.01:
            try:
                data = MINERAL_DATA["temp_25"][mineral][citation]
                return data["alpha"]
            except KeyError:
                # If 25C specific data missing, fall through to equation calculation
                pass

        # Calculate based on Temperature Equations
        try:
            data = MINERAL_DATA["temp_non_25"][mineral][citation]
            w = data["w"]
            x = data["x"]
            eqn_type = data["eqn"]

            if eqn_type == "T":
                # Equation: 1000 ln(alpha) = w * (10^3/T) + x
                val = (w * (1000 / temp_k)) + x
                return math.exp(val / 1000)
            
            elif eqn_type == "TSqr":
                # Equation: 1000 ln(alpha) = w * (10^6/T^2) + x
                val = (w * (1e6 / (temp_k**2))) + x
                return math.exp(val / 1000)
                
        except KeyError:
            raise ValueError(f"Mineral '{mineral}' with citation '{citation}' not found in database.")

    @staticmethod
    def calculate_slope(y_std_true, y_nbs18_true, x_std_meas, x_nbs18_meas, alpha_calcite, alpha_sample):
        """
        Calculates Slope (m) for the normalization equation y = mx + b
        """
        # Convert delta values to ratios logic implicitly handled here via simplification
        # Using JS logic: 
        # m = [((y2/1000+1)*(ac/as) - 1) - ((y1/1000+1)*(ac/as) - 1)] / (x2/1000 - x1/1000)
        
        term_std = ((y_std_true / 1000 + 1) * (alpha_calcite / alpha_sample)) - 1
        term_nbs18 = ((y_nbs18_true / 1000 + 1) * (alpha_calcite / alpha_sample)) - 1
        
        denominator = (x_std_meas / 1000) - (x_nbs18_meas / 1000)
        
        if denominator == 0:
            return 0 # Avoid division by zero
            
        m = (term_std - term_nbs18) / denominator
        return m

    @staticmethod
    def calculate_intercept(y_true, x_meas, m, alpha_calcite, alpha_sample):
        """
        Calculates Intercept (b) for the normalization equation y = mx + b
        """
        # b = [(y_true/1000 + 1) * (ac/as) - 1] - m * (x_meas/1000)
        # Result is multiplied by 1000 to get per mil
        
        term_y = ((y_true / 1000 + 1) * (alpha_calcite / alpha_sample)) - 1
        b_val = term_y - (m * (x_meas / 1000))
        
        return b_val * 1000

    @classmethod
    def process_sample(cls, 
                       meas_nbs18, 
                       meas_std_2, 
                       meas_sample, 
                       temp_c, 
                       mineral, 
                       citation, 
                       std_2_name="NBS19", # or "IAEA-603"
                       user_alpha=None):
        """
        Main function to call from Excel/Pandas.
        Returns a dictionary with VPDB and VSMOW values.
        """
        
        # 1. Get Known Standard Values (True Values)
        nbs18_vpdb = STD_VALUES["NBS18"]["VPDB"]
        nbs18_vsmow = STD_VALUES["NBS18"]["VSMOW"]
        
        std2_vpdb = STD_VALUES[std_2_name]["VPDB"]
        std2_vsmow = STD_VALUES[std_2_name]["VSMOW"]

        # 2. Calculate Alphas
        alpha_calcite = cls.get_calcite_acid_fractionation(temp_c)
        alpha_sample = cls.get_mineral_alpha(mineral, citation, temp_c, user_alpha)

        # 3. VPDB Calculation
        # Calculate Slope m
        m_vpdb = cls.calculate_slope(std2_vpdb, nbs18_vpdb, meas_std_2, meas_nbs18, alpha_calcite, alpha_sample)
        # Calculate Intercept b (using Standard 2 as the anchor point)
        b_vpdb = cls.calculate_intercept(std2_vpdb, meas_std_2, m_vpdb, alpha_calcite, alpha_sample)
        
        # Calculate final Sample VPDB
        sample_vpdb = (m_vpdb * meas_sample) + b_vpdb

        # 4. VSMOW Calculation
        # Calculate Slope m
        m_vsmow = cls.calculate_slope(std2_vsmow, nbs18_vsmow, meas_std_2, meas_nbs18, alpha_calcite, alpha_sample)
        # Calculate Intercept b
        b_vsmow = cls.calculate_intercept(std2_vsmow, meas_std_2, m_vsmow, alpha_calcite, alpha_sample)
        
        # Calculate final Sample VSMOW
        sample_vsmow = (m_vsmow * meas_sample) + b_vsmow

        return {
            "d18O_VPDB": round(sample_vpdb, 2),
            "d18O_VSMOW": round(sample_vsmow, 2),
            "alpha_sample": round(alpha_sample, 5),
            "alpha_calcite": round(alpha_calcite, 5),
            "slope_vpdb": round(m_vpdb, 5),
            "intercept_vpdb": round(b_vpdb, 5)
        }
    
"""
Sample Usage

# Import this module
import oxygen_isotope_calculator as iso

# ==========================================
# EXAMPLE USAGE
# ==========================================

# 1. Setup your inputs (These would come from your Excel row)
inputs = {
    "meas_nbs18": -23.15,      # Measured NBS18
    "meas_std_2": -2.10,       # Measured NBS19 or IAEA-603
    "meas_sample": -5.50,      # Measured Sample
    "temp_c": 70.0,            # Acid Temp
    "mineral": "Calcite",      # Mineral Type
    "citation": "Kim et al. (2007)", 
    "std_2_name": "NBS19"      # Which second standard did you use?
}

# 2. Run the calculation
try:
    result = iso.OxygenIsotopeCalculator.process_sample(
        meas_nbs18=inputs["meas_nbs18"],
        meas_std_2=inputs["meas_std_2"],
        meas_sample=inputs["meas_sample"],
        temp_c=inputs["temp_c"],
        mineral=inputs["mineral"],
        citation=inputs["citation"],
        std_2_name=inputs["std_2_name"]
    )

    print("--- Calculation Results ---")
    print(f"Sample d18O (VPDB):  {result['d18O_VPDB']} permil")
    print(f"Sample d18O (VSMOW): {result['d18O_VSMOW']} permil")
    print(f"Alpha Used:          {result['alpha_sample']}")

except Exception as e:
    print(f"Error: {e}")



"""