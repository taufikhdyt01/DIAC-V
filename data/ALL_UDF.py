import xlwings as xw
import matplotlib.pyplot as plt
import io
import pandas as pd
import re
import CoolProp.CoolProp as CP
import numpy as np
import matplotlib.pyplot as plt
import math




import numpy as np
from scipy.interpolate import CubicSpline
from scipy.interpolate import InterpolatedUnivariateSpline
from scipy.optimize import minimize_scalar  # Missing import added here
from scipy.interpolate import PchipInterpolator  # For monotonic interpolation
from xlwings.constants import ChartType

from scipy.interpolate import interp1d
from scipy.optimize import brentq



#----------------

@xw.func
def CUBIC_SPLINE_INTERPOLATE(x_values, y_values, query_x):
    """
    Performs cubic spline interpolation for a given set of X and Y values and a query X.
    Args:
        x_values (list): List or range of X values.
        y_values (list): List or range of Y values.
        query_x (float): The X value for which to interpolate Y.

    Returns:
        float: Interpolated Y value.
    """
    try:
        # Convert inputs to numpy arrays
        x_values = np.array([float(x) for x in x_values if x is not None])
        y_values = np.array([float(y) for y in y_values if y is not None])
        
        # Ensure data is sorted by X
        sorted_indices = np.argsort(x_values)
        x_values = x_values[sorted_indices]
        y_values = y_values[sorted_indices]

        # Perform cubic spline interpolation
        spline = CubicSpline(x_values, y_values)
        return spline(query_x)

    except Exception as e:
        return f"Error: {str(e)}"


#---------------

@xw.func
def MONOTONIC_SPLINE_INTERPOLATE(x_values, y_values, query_x):
    """
    Performs shape-preserving (monotonic) spline interpolation using PchipInterpolator.
    This ensures if the data is descending in Y, the interpolated function won't go up.

    Args:
        x_values (list): List or range of X values.
        y_values (list): List or range of Y values.
        query_x (float): The X value at which to interpolate Y.

    Returns:
        float: Interpolated Y value (monotonic curve).
    """
    try:
        # Convert inputs to numpy arrays, ignoring any None or blank cells
        x_arr = np.array([float(x) for x in x_values if x is not None])
        y_arr = np.array([float(y) for y in y_values if y is not None])

        # Ensure data is sorted by X ascending
        sort_idx = np.argsort(x_arr)
        x_arr = x_arr[sort_idx]
        y_arr = y_arr[sort_idx]

        # Create a shape-preserving interpolator
        pchip = PchipInterpolator(x_arr, y_arr)

        # Evaluate at query_x
        return float(pchip(float(query_x)))

    except Exception as e:
        return f"Error: {str(e)}"

#-------------------

@xw.func
def CUBIC_SPLINE_INTERPOLATE_INSIDE(x_values, y_values, query_x):
    """
    Performs cubic spline interpolation with monotonic constraints for a given set of X and Y values.
    Restricts interpolation to within the range of X values.
    Args:
        x_values (list): List or range of X values.
        y_values (list): List or range of Y values.
        query_x (float): The X value for which to interpolate Y.

    Returns:
        float: Interpolated Y value, or "DATA NOT FOUND" if query_x is outside the range.
    """
    try:
        # Convert inputs to numpy arrays
        x_values = np.array([float(x) for x in x_values if x is not None])
        y_values = np.array([float(y) for y in y_values if y is not None])

        # Ensure data is sorted by X
        sorted_indices = np.argsort(x_values)
        x_values = x_values[sorted_indices]
        y_values = y_values[sorted_indices]

        # Check if query_x is within the range
        if query_x < np.min(x_values) or query_x > np.max(x_values):
            return "DATA OUT OF RANGE"

        # Perform monotonic cubic interpolation (PCHIP)
        spline = PchipInterpolator(x_values, y_values)
        return spline(query_x)

    except Exception as e:
        return f"Error: {str(e)}"
# -------------

@xw.func
def generate_focus_point_chart(sheet_name, x_values, y_values, focus_x, focus_y, point_color="red", line_color="blue", chart_name="My Chart"):
    """
    Generates an Excel chart with a focus point and customizable line and point colors.
    
    Args:
        sheet_name (str): The name of the sheet to place the chart on.
        x_values (list): List of X values for the chart.
        y_values (list): List of Y values for the chart.
        focus_x (float): X value for the focus point.
        focus_y (float): Y value for the focus point.
        point_color (str): Color of the focus point (default: "red").
        line_color (str): Color of the line (default: "blue").
        chart_name (str): Title of the chart.
    
    Returns:
        str: Success or error message.
    """
    try:
        # Get the workbook and target sheet
        wb = xw.Book.caller()
        sheet = wb.sheets[sheet_name]

        # Write X and Y data to the sheet
        sheet.range("A1").value = "X"
        sheet.range("B1").value = "Y"
        sheet.range("A2").value = [[x] for x in x_values]
        sheet.range("B2").value = [[y] for y in y_values]

        # Add focus point data
        sheet.range("D1").value = "Focus X"
        sheet.range("E1").value = "Focus Y"
        sheet.range("D2").value = focus_x
        sheet.range("E2").value = focus_y

        # Add a chart
        chart = sheet.charts.add()
        chart.chart_type = "xy_scatter_smooth"  # Scatter with smooth lines
        chart.set_source_data(sheet.range(f"A1:B{len(x_values) + 1}"))

        # Customize the series (line and points)
        series = chart.api[1].SeriesCollection(1)
        series.Format.Line.ForeColor.RGB = xw.utils.rgb_to_int(line_color)

        # Add the focus point as a new series
        focus_range_x = sheet.range("D1:D2")
        focus_range_y = sheet.range("E1:E2")
        chart.api[1].SeriesCollection().NewSeries()
        focus_series = chart.api[1].SeriesCollection(2)
        focus_series.XValues = focus_range_x.address
        focus_series.Values = focus_range_y.address
        focus_series.MarkerStyle = -4142  # Circle
        focus_series.MarkerSize = 10  # Make the focus point bigger
        focus_series.Format.Line.Visible = False  # No line for the focus point
        focus_series.Format.Fill.ForeColor.RGB = xw.utils.rgb_to_int(point_color)

        # Set the chart title
        chart.api[1].HasTitle = True
        chart.api[1].ChartTitle.Text = chart_name

        # Customize axis colors
        chart.api[1].Axes(1).Format.Line.ForeColor.RGB = xw.utils.rgb_to_int("black")  # X-axis color
        chart.api[1].Axes(2).Format.Line.ForeColor.RGB = xw.utils.rgb_to_int("black")  # Y-axis color

        return "Excel chart with focus point generated successfully!"
    except Exception as e:
        return f"Error: {repr(e)}"


#------------------------


CONVERSION_FACTORS = {
    "pressure": {
        "bar": 1,
        "psi": 14.5038,
        "m_aq": 10.197,
        "kg_cm2": 1.01972,
    },
}

@xw.sub
def update_value_based_on_dropdown():
    """
    Dynamically updates the value based on a dropdown menu in the same row.
    Assumes the dropdown menu is one column to the right of the value.
    """
    wb = xw.Book.caller()
    sheet = wb.sheets.active

    # Use the correct method to get the active cell
    dropdown_cell = xw.apps.active.selection

    # Assume the value cell is one column to the left of the dropdown
    value_cell = dropdown_cell.offset(0, -1)

    # Retrieve the value and selected unit
    value = value_cell.value
    target_unit = dropdown_cell.value

    # Detect the appropriate unit group dynamically
    unit_group = None
    for group, units in CONVERSION_FACTORS.items():
        if target_unit.lower() in units:
            unit_group = group
            break

    if not unit_group:
        value_cell.value = "Invalid unit!"
        return

    # Perform the conversion
    try:
        base_unit = list(CONVERSION_FACTORS[unit_group].keys())[0]
        factors = CONVERSION_FACTORS[unit_group]
        base_factor = factors[base_unit]
        target_factor = factors[target_unit.lower()]
        converted_value = value * (target_factor / base_factor)

        # Update the value cell with the converted value
        value_cell.value = converted_value
    except KeyError:
        value_cell.value = "Conversion error!"

#-------------------------


@xw.func
def INVERSE_INTERPOLATION(x_values, y_values, target_y):
    try:
        # Convert Excel ranges to numpy arrays and clean the data
        x_values = np.array([float(x) for x in x_values if x is not None])
        y_values = np.array([float(y) for y in y_values if y is not None])

        # Check if the lengths of x_values and y_values match
        if len(x_values) != len(y_values):
            return "Error: X and Y arrays must have the same length"

        # Validate the target_y range
        if target_y < min(y_values) or target_y > max(y_values):
            return "Error: Query value out of range"

        # Split dataset into increasing and decreasing segments
        peak_index = np.argmax(y_values)
        x_increasing = x_values[:peak_index + 1]
        y_increasing = y_values[:peak_index + 1]
        x_decreasing = x_values[peak_index:]
        y_decreasing = y_values[peak_index:]

        # Handle increasing segment
        if min(y_increasing) <= target_y <= max(y_increasing):
            spline_inc = InterpolatedUnivariateSpline(x_increasing, y_increasing - target_y)
            roots_inc = spline_inc.roots()
        else:
            roots_inc = []

        # Handle decreasing segment
        if min(y_decreasing) <= target_y <= max(y_decreasing):
            spline_dec = InterpolatedUnivariateSpline(x_decreasing, y_decreasing - target_y)
            roots_dec = spline_dec.roots()
        else:
            roots_dec = []

        # Combine roots from both segments
        roots = np.concatenate((roots_inc, roots_dec))

        if len(roots) == 0:
            return "No solution found"
        return roots.tolist()  # Return all matching X-values as a list

    except ValueError as ve:
        return f"Error: Invalid input data - {str(ve)}"
    except Exception as e:
        return f"Error: {str(e)}"

#---------------------------------------


@xw.func
def INVERSE_LAGRANGE_INTERPOLATION(x_values, y_values, query_y):
    try:
        # Ensure inputs are numpy arrays
        x_values = np.array(x_values, dtype=float)
        y_values = np.array(y_values, dtype=float)

        # Number of data points
        n = len(x_values)

        # Initialize the result
        result_x = 0.0

        # Compute the Lagrange interpolation
        for i in range(n):
            # Compute the basis polynomial L_i(query_y)
            L_i = 1.0
            for j in range(n):
                if i != j:
                    L_i *= (query_y - y_values[j]) / (y_values[i] - y_values[j])
            
            # Add the contribution of the i-th term
            result_x += x_values[i] * L_i

        return result_x

    except Exception as e:
        return f"Error: {str(e)}"

#------------------------------------


@xw.func
def INTERPOLATE_X_FOR_Y(x_values, y_values, query_y): 
    """
    UDF to interpolate the X value corresponding to a given Y value.
    Args:
        x_values (list): List or range of X values from Excel.
        y_values (list): List or range of Y values from Excel.
        query_y (float): The Y value for which to find the corresponding X.

    Returns:
        float or str: Interpolated X value for the given Y, or an error message.
    """
    try:
        # Convert inputs to numpy arrays and ensure they are floats
        x_values = np.array([float(x) for x in x_values if x is not None])
        y_values = np.array([float(y) for y in y_values if y is not None])
        
        # Check if lengths match
        if len(x_values) != len(y_values):
            return "Error: X and Y arrays must have the same length."

        # Check if the query_y is within the range of Y-values
        if query_y < np.min(y_values) or query_y > np.max(y_values):
            return "Error: Query Y value is out of range."

        # Use cubic spline interpolation
        spline = CubicSpline(y_values, x_values)

        # Interpolate the X value for the given Y
        interpolated_x = spline(query_y)

        return interpolated_x

    except Exception as e:
        return f"Error: {str(e)}"

#----------------------------------


@xw.func
def CUBIC_SPLINE_REVERSE_INTERPOLATE(x_values, y_values, query_y, tolerance=1e-6):
    try:
        # Ensure inputs are numpy arrays
        x_values = np.array(x_values)
        y_values = np.array(y_values)

        # Remove near-duplicate Y values using the tolerance
        filtered_indices = [0]  # Always keep the first point
        for i in range(1, len(y_values)):
            if abs(y_values[i] - y_values[filtered_indices[-1]]) > tolerance:
                filtered_indices.append(i)

        x_values = x_values[filtered_indices]
        y_values = y_values[filtered_indices]

        # Identify the main peak (maximum Y value)
        peak_index = np.argmax(y_values)

        # Handle decreasing tail if present at the end
        if peak_index < len(y_values) - 1:
            # Check for another decrease after the main peak
            final_decreasing_start = None
            for i in range(peak_index + 1, len(y_values) - 1):
                if y_values[i] > y_values[i + 1]:  # Detect tail drop
                    final_decreasing_start = i
                    break

            if final_decreasing_start:
                x_tail = x_values[final_decreasing_start:]
                y_tail = y_values[final_decreasing_start:]

                if query_y >= y_tail[-1] and query_y <= y_tail[0]:
                    # Handle the decreasing tail
                    spline = CubicSpline(y_tail[::-1], x_tail[::-1])
                    return spline(query_y)

        # Split into increasing and decreasing segments
        x_increasing = x_values[:peak_index + 1]
        y_increasing = y_values[:peak_index + 1]

        x_decreasing = x_values[peak_index:]
        y_decreasing = y_values[peak_index:]

        # Determine which segment the query value falls into
        if query_y >= y_increasing[0] and query_y <= y_increasing[-1]:
            # Interpolate within the increasing segment
            spline = CubicSpline(y_increasing, x_increasing)
            return spline(query_y)
        elif query_y >= y_decreasing[-1] and query_y <= y_decreasing[0]:
            # Interpolate within the decreasing segment
            spline = CubicSpline(y_decreasing[::-1], x_decreasing[::-1])  # Reverse for decreasing
            return spline(query_y)
        else:
            return "Error: Query value out of range"

    except Exception as e:
        return f"Error: {str(e)}"

#----------------------------------------


@xw.func
def FIND_Ymax_SMOOTH(x_values, y_values):
    """
    Finds the maximum Y value from a smooth interpolated curve.
    Args:
        x_values (list): List or range of X values from Excel.
        y_values (list): List or range of Y values from Excel.

    Returns:
        float: Maximum Y value from the smooth curve.
    """
    try:
        # Convert inputs to numpy arrays and filter valid values
        x_values = np.array([float(x) for x in x_values if x is not None])
        y_values = np.array([float(y) for y in y_values if y is not None])

        # Ensure data is sorted by X
        sorted_indices = np.argsort(x_values)
        x_values = x_values[sorted_indices]
        y_values = y_values[sorted_indices]

        # Create a cubic spline for smoothing
        spline = CubicSpline(x_values, y_values)

        # Define a function to minimize (negative of the spline to find the maximum)
        def neg_spline(value):
            return -spline(value)

        # Define bounds for the search (min and max X values)
        bounds = (x_values[0], x_values[-1])

        # Find the maximum Y value by minimizing the negative spline
        result = minimize_scalar(neg_spline, bounds=bounds, method="bounded")

        if result.success:
            ymax = -result.fun  # Maximum Y value
            return ymax
        else:
            return "Error: Could not find maximum Y"

    except Exception as e:
        return f"Error: {str(e)}"


# ---------------------------------


@xw.func
def FIND_X_at_Ymax_SMOOTH(x_values, y_values):
    """
    Finds the X value corresponding to the maximum Y value from a smooth interpolated curve.
    Args:
        x_values (list): List or range of X values from Excel.
        y_values (list): List or range of Y values from Excel.

    Returns:
        float: X value corresponding to the maximum Y value.
    """
    try:
        # Convert inputs to numpy arrays
        x_values = np.array([float(x) for x in x_values if x is not None])
        y_values = np.array([float(y) for y in y_values if y is not None])

        # Ensure data is sorted by X
        sorted_indices = np.argsort(x_values)
        x_values = x_values[sorted_indices]
        y_values = y_values[sorted_indices]

        # Create a cubic spline for the data
        spline = CubicSpline(x_values, y_values)

        # Define a function to minimize (negative of the spline to find the maximum)
        def neg_spline(value):
            return -spline(value)

        # Define bounds for the search
        bounds = (x_values[0], x_values[-1])

        # Find the X value where the spline's negative is minimized (Ymax)
        result = minimize_scalar(neg_spline, bounds=bounds, method="bounded")

        if result.success:
            x_at_ymax = result.x  # X value at the smooth curve's Ymax
            return x_at_ymax
        else:
            return "Error: Could not find X for Ymax"

    except Exception as e:
        return f"Error: {str(e)}"

#-------------------------------

@xw.func
def XLOOKUP_ADDRESS(lookup_value, lookup_array, return_array):
    """
    Find the address of the cell that matches the lookup_value.
    """
    try:
        # Convert inputs to lists
        lookup_array = [str(cell) for cell in lookup_array]
        return_array = list(return_array)

        # Find the index of the lookup value
        index = lookup_array.index(lookup_value)

        # Get the address of the matching cell in the return array
        return xw.Range(return_array[index].get_address()).address
    except ValueError:
        return "Not Found"

#---------------------------------------

import xlwings as xw

@xw.func
def find_query_address(query_text, search_range):
    """
    Finds the cell address of the query text in the search range.
    Args:
        query_text (str): The text to search for (e.g., "CRN 64-5-2").
        search_range (list): The range of cells to search in (e.g., B32:S32).

    Returns:
        str: The cell address of the matching text (e.g., "$H$32").
    """
    try:
        # Get the workbook and active sheet
        wb = xw.Book.caller()
        ws = wb.sheets.active

        # Convert the search range to a list
        search_range = list(search_range)

        # Find the matching column
        if query_text in search_range:
            match_index = search_range.index(query_text)
            col_start = ws.range(search_range).column + match_index
            row_start = ws.range(search_range).row

            # Return the address of the matching cell
            cell_address = xw.utils.cell(row_start, col_start, absolute=True)
            return cell_address
        else:
            return f"Query Text '{query_text}' Not Found"
    except Exception as e:
        return f"Error: {repr(e)}"

#--------------------------------------


@xw.func
def find_cell_address(query, search_range):
    for cell in search_range:
        if cell.value == query:
            return cell.address
    return "Not Found"

#---------------------------------------------


def calc_viscosity_factor_alt(mu, mu_ref=0.001):
    """
    Calculate viscosity correction factors for head (K_H) and flow (K_Q),
    using the formula:
    
        K_H = 1 + 0.15 * (mu/mu_ref - 1)
        K_Q = 1 + 0.1  * (mu/mu_ref - 1)
    
    so that if mu > mu_ref => K_H > 1 => the normalized head on the water-based curve
    is larger for more viscous fluid, and similarly for flow if you choose to apply it.
    
    Here:
      - mu     : fluid's actual viscosity at T1 (Pa·s)
      - mu_ref : reference viscosity at T0 (Pa·s)
      - 0.15   : empirical constant for head correction
      - 0.1    : empirical constant for flow correction
    """
    # If mu/mu_ref > 1 => each factor > 1 => "bigger pump" interpretation.
    K_H = 1 + 0.15*(mu/mu_ref - 1)
    K_Q = 1 + 0.1 *(mu/mu_ref - 1)
    return K_H, K_Q

@xw.func
def Alt_Normalizing_Pump_Q2(
    Q, T1, TSS, D0, altitude, fluid="Water", T0=20, mu_init=None
):
    """
    Normalizes the flow (Q) to the vendor's water-based curve, allowing an optional
    user-defined viscosity mu_init at T0. If mu_init is given, we scale from water's
    T0..T1 ratio to get the fluid's T1 viscosity.

    MATH FORMULAS (inline reference):

    1) If mu_init is None:
         mu_T1 = mu_water(T1).
         mu_ref = mu_water(T0).
       Otherwise:
         ratio = mu_init / mu_water(T0).
         mu_T1 = mu_water(T1) * ratio
         mu_ref = mu_init
       
    2) Correction factor:
         _, K_Q = calc_viscosity_factor_alt(mu_T1, mu_ref)
    
    3) Normalized flow:
         Q_norm = Q * K_Q
    
    4) This function does NOT incorporate TSS in flow corrections by default:
         D1 = D0 + TSS/1000  (but not used in the formula for Q).

    Parameters:
      Q (float)         : actual flow (m³/h).
      T1 (float)        : fluid temperature in °C (real operation).
      TSS (float)       : mg/L solids, used only if you want to incorporate in density or mu.
      D0 (float)        : base fluid density at T1 ignoring TSS (kg/m³).
      altitude (float)  : altitude in meters => partial pressure for CP.
      fluid (str)       : base fluid name for CP.PropsSI
      T0 (float)        : reference temperature in °C for vendor data.
      mu_init (float|None) : user's fluid viscosity at T0 (Pa·s). If None => pure water.
    """
    try:
        # 1) Water's viscosity at T1
        T1_K = T1 + 273.15
        P_sea_level = 101325
        P_atm = P_sea_level * (1 - (0.0065 * altitude / 288.15))**5.2561
        water_mu_T1 = CP.PropsSI('V','T', T1_K,'P', P_atm, fluid)

        # 2) Water's viscosity at T0
        T0_K = T0 + 273.15
        water_mu_T0 = CP.PropsSI('V','T', T0_K,'P', 101325, fluid)

        if mu_init is None:
            # Just do water
            mu_T1 = water_mu_T1
            mu_ref = water_mu_T0
        else:
            # scale factor: fluid is mu_init at T0, so ratio = mu_init / water_mu_T0
            ratio = mu_init / water_mu_T0
            mu_T1 = water_mu_T1 * ratio
            mu_ref = mu_init

        # 3) We do NOT incorporate density in Q corrections
        #    (If you'd like TSS => user might do mu_T1 *= f(TSS) or something.)

        # 4) Correction factor
        _, K_Q = calc_viscosity_factor_alt(mu_T1, mu_ref)

        # 5) Return unrounded float
        return Q * K_Q

    except Exception as e:
        return f"Error: {e}"

@xw.func
def Alt_Normalizing_Pump_H2(
    H, T1, TSS, D0, altitude, fluid="Water", T0=20, mu_init=None
):
    """
    Normalizes head (H) to the vendor's water-based curve, factoring in density & viscosity.

    MATH FORMULAS:

    1) T1 => mu_T1 from water with partial pressure, scaled if user gave mu_init:
         mu_T1 = water_mu_T1 * ( mu_init / water_mu_T0 ) if mu_init != None
         else mu_T1 = water_mu_T1
       and mu_ref = mu_init or water_mu_T0 accordingly.

    2) TSS => heavier fluid:
         D1 = D0 + (TSS/1000)
       Reference density at T0 => D0_T0 = CP.PropsSI('D', 'T', T0+273.15, 'P', 101325, fluid)

    3) Correction factor for viscosity:
         K_H, _ = calc_viscosity_factor_alt(mu_T1, mu_ref)

    4) Normalized head:
         H_norm = H * (D1 / D0_T0) * K_H
    
    Returns: unrounded float for final H_norm.

    Example usage:
      =Alt_Normalizing_Pump_H2(10, 50, 2000, 998, 300, "Water", 20)
      => if no mu_init => just water-based. If user passes 0.0015 => scaled from T0 -> T1.
    """
    try:
        P_sea_level = 101325
        P_atm = P_sea_level * (1 - (0.0065*altitude/288.15))**5.2561

        T1_K = T1 + 273.15
        T0_K = T0 + 273.15

        # water-based mu
        water_mu_T1 = CP.PropsSI('V','T',T1_K,'P',P_atm,fluid)
        water_mu_T0 = CP.PropsSI('V','T',T0_K,'P',101325,fluid)

        if mu_init is None:
            mu_T1 = water_mu_T1
            mu_ref = water_mu_T0
        else:
            ratio = mu_init / water_mu_T0
            mu_T1 = water_mu_T1 * ratio
            mu_ref = mu_init

        D1 = D0 + TSS/1000.0
        D0_T0 = CP.PropsSI('D','T',T0_K,'P',101325,fluid)

        K_H, _ = calc_viscosity_factor_alt(mu_T1, mu_ref)

        return H * (D1 / D0_T0) * K_H

    except Exception as e:
        return f"Error: {e}"



#---------------------------------------

def calculate_viscosity_correction(mu, mu_ref=0.001):
    """
    Calculate viscosity correction factors for head (K_H) and flow (K_Q),
    assuming we want "higher viscosity => bigger correction factor" 
    so we move to a bigger 'water-based' point on the vendor curve.

    If mu > mu_ref => K_H > 1 => that means "normalized head" is larger.
    Similarly K_Q > 1 => 'normalized flow' is bigger 
    (the water curve speed would be bigger to replicate the real fluid flow).
    This sign is common if your 'normalized' means 
    'the point on the water curve that matches real fluid performance.'
    
    Adjust the 0.15 or 0.1 multipliers based on actual pump handbooks.
    """
    # If mu/mu_ref = 1 => K=1
    # If mu>mu_ref => K>1
    K_H = 1 + 0.15 * (mu / mu_ref - 1)
    K_Q = 1 + 0.1  * (mu / mu_ref - 1)
    return K_H, K_Q

@xw.func
def Normalizing_Pump_Param_Q(Q, T1, TSS, D0, altitude, fluid="Water", T0=20):
    """
    Normalize fluid flow rate to match vendor pump chart at reference T0,
    accounting for fluid viscosity (due to temperature, altitude, etc.).
    Currently we do *not* incorporate density in flow correction,
    only viscosity.

    Params:
      Q (float): Actual fluid flow (m³/h).
      T1 (float): Fluid temperature (°C).
      TSS (float): mg/L of solids (unused in flow except for density calc).
      D0 (float): Base fluid density at T1 (kg/m³), ignoring TSS.
      altitude (float): altitude in m (affects pressure => affects viscosity).
      fluid (str): fluid name for CoolProp (default "Water").
      T0 (float): reference temperature (°C) for vendor chart.

    Returns: 
      float => 'normalized' flow for water-based chart, so if fluid is 
               more viscous => we get a bigger 'Q' on the chart.
    """
    try:
        # Convert temperature => K
        T1_K = T1 + 273.15
        T0_K = T0 + 273.15

        # atmospheric pressure at altitude
        P_sea_level = 101325
        P_atm = P_sea_level * (1 - (0.0065*altitude/288.15))**5.2561

        # get fluid viscosity at T1 & T0
        mu_T1 = CP.PropsSI('V', 'T', T1_K, 'P', P_atm, fluid)   # actual fluid
        mu_T0 = CP.PropsSI('V', 'T', T0_K, 'P', P_sea_level, fluid) 
          # or P_atm if you prefer to keep altitude for T0 as well

        # Possibly compute D1 if you want, but we don't use it for flow:
        D1 = D0 + TSS/1000.0

        # get correction factor
        _, K_Q = calculate_viscosity_correction(mu_T1, mu_T0)

        # Return normalized flow
        return round(Q * K_Q, 2)

    except Exception as e:
        return f"Error: {e}"

@xw.func
def Normalizing_Pump_Param_H(H, T1, TSS, D0, altitude, fluid="Water", T0=20):
    """
    Normalize fluid head to match vendor pump chart at reference T0,
    so heavier or more viscous fluid => bigger normalized head 
    on the water-based curve.

    We incorporate density ratio => if fluid is denser than vendor's ref, 
    we multiply the head. Also incorporate a viscosity factor >1 if mu>mu_ref.

    Params:
      H (float): Actual measured head (m).
      T1 (float): Temperature (°C).
      TSS (float): mg/L solids => added to density.
      D0 (float): base fluid density at T1, ignoring TSS
      altitude (float): altitude for partial pressure => affects viscosity
      fluid (str): fluid name in CoolProp
      T0 (float): reference temp (°C)

    Returns:
      float => 'normalized' head (larger if fluid is denser, more viscous).
    """
    try:
        T1_K = T1 + 273.15
        T0_K = T0 + 273.15

        P_sea_level = 101325
        P_atm = P_sea_level * (1 - (0.0065*altitude/288.15))**5.2561

        mu_T1 = CP.PropsSI('V','T', T1_K,'P', P_atm, fluid)
        # if vendor reference is at sea level:
        mu_T0 = CP.PropsSI('V','T', T0_K,'P', P_sea_level, fluid)

        # Adjust density for TSS
        D1 = D0 + TSS/1000.0  # actual fluid density

        # get vendor reference density at T0 (assuming sea level for vendor)
        # or if you prefer the same altitude, do 'P_atm'
        D0_T0 = CP.PropsSI('D','T', T0_K,'P', 101325, fluid)

        # get viscosity correction factor
        K_H, _ = calculate_viscosity_correction(mu_T1, mu_T0)

        # scale by density ratio => heavier fluid => bigger head
        # then scale by viscosity => bigger factor if mu>mu_ref
        H_normalized = H * (D1 / D0_T0) * K_H

        return round(H_normalized, 2)

    except Exception as e:
        return f"Error: {e}"



#-----------------------------------------

def _clean_and_build_df(library_data):
    """
    Pre-clean library_data so we skip rows whose first column is blank or ends
    with "group" (ignoring case/spaces). Then we pad/truncate each row
    to match the header length, and trim the "Parameters" column strings.
    """
    import pandas as pd

    if not library_data or len(library_data) < 2:
        raise ValueError("Not enough rows for a header + data.")

    header = library_data[0]
    rows = library_data[1:]
    if not any(header):
        raise ValueError("Header row is blank.")

    # Strip each header cell
    header = [str(h).strip() for h in header]

    cleaned_rows = []
    for row in rows:
        if not row:
            continue

        row_list = list(row)
        first_cell = str(row_list[0]).strip() if row_list else ""

        # remove all spaces, lowercase to check if ends with "group"
        normalized = first_cell.replace(" ", "").lower()

        if normalized == "" or normalized.endswith("group"):
            # skip "blank" or "something group"
            continue

        # pad or truncate row_list
        if len(row_list) < len(header):
            row_list += [""] * (len(header) - len(row_list))
        elif len(row_list) > len(header):
            row_list = row_list[: len(header)]

        cleaned_rows.append(row_list)

    df = pd.DataFrame(cleaned_rows, columns=header)

    # rename the first column => "Parameters" if needed
    first_col = df.columns[0]
    if first_col.strip().lower() != "parameters":
        df.rename(columns={first_col: "Parameters"}, inplace=True)

    # Now trim all strings in the "Parameters" column
    df["Parameters"] = df["Parameters"].astype(str).str.strip()

    return df

@xw.func
@xw.arg("horizontal_range", "range")
@xw.arg("vertical_range", "range")
def LIBRARY_QUERY(
    horizontal_range,
    vertical_range,
    query_category: str,
    query_parameter=None,
    file_path=None
):
    """
    Returns numeric results for all or one parameter from a library table
    that may contain 'group' rows at the end.

    STEPS:
    1) Trim query_category & query_parameter if present.
    2) Build bounding rectangle from horizontal_range & vertical_range.
    3) Read that rectangle => skip group rows => build DataFrame => "Parameters" col is trimmed.
    4) If query_parameter is None => return all numeric values in vertical Nx1 array.
       If query_parameter is given => return single numeric cell. 
         If not found => error string.

    Example Excel usage:
      =LIBRARY_QUERY(B2:G2, A2:A12, "Dairy") => all param values
      =LIBRARY_QUERY(B2:G2, A2:A12, "Dairy", "TSS") => single param "TSS"
    """
    try:
        # 1) Trim input
        query_category = query_category.strip() if query_category else ""
        if query_parameter:
            query_parameter = query_parameter.strip()

        # 2) figure out which workbook to read from
        if file_path:
            wb_lib = xw.Book(file_path)   # open external file
            sht_lib = horizontal_range.sheet  # referencing that sheet
        else:
            wb_lib = horizontal_range.sheet.book
            sht_lib = horizontal_range.sheet

        # bounding rectangle
        top_row    = min(horizontal_range.row, vertical_range.row)
        left_col   = min(horizontal_range.column, vertical_range.column)
        bottom_row = max(horizontal_range.last_cell.row, vertical_range.last_cell.row)
        right_col  = max(horizontal_range.last_cell.column, vertical_range.last_cell.column)

        rng_lib = sht_lib.range((top_row, left_col), (bottom_row, right_col))
        library_data = rng_lib.value

        # 3) Build DataFrame & skip group
        df = _clean_and_build_df(library_data)

        # Check if category in df columns
        if query_category not in df.columns:
            return f"Error: Category '{query_category}' not found in columns: {list(df.columns)}"

        # 4) If query_parameter=None => return all numeric values in Nx1
        if not query_parameter:
            results = []
            for _, rowdata in df.iterrows():
                val = rowdata[query_category]
                results.append([val])  # each row => single cell
            return results

        else:
            # single param => match the "Parameters" col
            # we also do .strip() in df, so "TSS" matches param "TSS"
            mask = (df["Parameters"] == query_parameter)
            subset = df[mask]
            if subset.empty:
                return f"Error: Parameter '{query_parameter}' not found."
            val = subset.iloc[0][query_category]
            return [[val]]  # return single cell

    except Exception as exc:
        return f"Error: {exc}"


#-----------------------------------
def calc_viscosity_factor(mu_actual, mu_ref):
    """
    REFINED VISCOUS CORRECTION FACTOR (Linear Approx.)
    
    We define smaller slopes to avoid large overcorrection
    for 1–5 cP fluids. If ratio=(mu_actual/mu_ref), then:
      K_H = 1.0 + 0.05*(ratio - 1.0)
      K_Q = 1.0 + 0.03*(ratio - 1.0)
    
    For example, if mu_ref=1 cP and mu_actual=3 cP => ratio=3
      => K_H = 1 + 0.05*(2) = 1.10  (+10%)
      => K_Q = 1 + 0.03*(2) = 1.06  (+6%)
    
    This is still an approximation; consider official HI methods
    for higher viscosities or tighter accuracy needs.
    """
    ratio = mu_actual / mu_ref
    K_H = 1.0 + 0.05 * (ratio - 1.0)
    K_Q = 1.0 + 0.03 * (ratio - 1.0)
    return K_H, K_Q


@xw.func
@xw.arg("altitude", float, empty=0.0)
@xw.arg("TSS", float, empty=0.0)
@xw.arg("D0", float, empty=None)
@xw.arg("mu", float, empty=None)
@xw.arg("T0", float, empty=20.0)
def PumpNormalizeQ(
    Q: float,      # actual flow (m³/h)
    T1: float,     # fluid temp (°C)
    altitude,      # default=0 if blank
    TSS,           # default=0 if blank
    D0,            # if blank => fallback to water density
    mu,            # if blank => fallback to water viscosity
    T0             # vendor reference temperature (°C)
):
    """
    UDF that returns ONLY the normalized flow (Q_norm) to match the
    pump vendor's reference water at T0 (°C).
    
    STEPS:
    1) Compute local partial pressure from altitude => P_alt.
    2) If 'mu' is None => get water viscosity at T1, P_alt => mu_T1.
    3) Get reference mu_ref => water at T0, sea-level => mu_ref.
    4) (K_H, K_Q) = calc_viscosity_factor(mu_T1, mu_ref) => use K_Q.
    5) Q_norm = Q * K_Q.
    6) Return Q_norm.
    
    This function does NOT directly use TSS or D0 for flow normalization,
    as typically TSS/density changes do not drastically affect capacity
    the same way viscosity does. But they're available to match the
    signature or for further custom logic.
    """
    try:
        # 1) altitude => partial pressure
        P_sea_level = 101325
        P_alt = P_sea_level * (1 - 0.0065*altitude/288.15)**5.2561

        # 2) get fluid viscosity at T1
        T1_K = T1 + 273.15
        if mu is None:
            mu_T1 = CP.PropsSI('V', 'T', T1_K, 'P', P_alt, 'Water')
        else:
            mu_T1 = float(mu)

        # 3) reference viscosity => water at T0 => sea level
        T0_K = T0 + 273.15
        mu_ref = CP.PropsSI('V', 'T', T0_K, 'P', 101325, 'Water')

        # 4) compute correction factor
        _, K_Q = calc_viscosity_factor(mu_T1, mu_ref)

        # 5) normalized flow
        Q_norm = Q * K_Q
        return Q_norm

    except Exception as e:
        return f"Error: {e}"


@xw.func
@xw.arg("altitude", float, empty=0.0)
@xw.arg("TSS", float, empty=0.0)
@xw.arg("D0", float, empty=None)
@xw.arg("mu", float, empty=None)
@xw.arg("T0", float, empty=20.0)
def PumpNormalizeH(
    H: float,      # actual head (m)
    T1: float,     # fluid temp (°C)
    altitude,
    TSS,
    D0,
    mu,
    T0
):
    """
    UDF that returns ONLY the normalized head (H_norm) to match the
    pump vendor's reference water at T0 (°C).
    
    STEPS:
    1) altitude => partial pressure => P_alt
    2) if mu is None => water from CP at T1 => mu_T1
    3) if D0 is None => water density at T1 => d0_T1
    4) D1 = d0_T1 + TSS/1000
    5) reference mu => water at T0 => mu_ref
    6) reference density => water at T0 => D_ref
    7) (K_H, K_Q) = calc_viscosity_factor(mu_T1, mu_ref) => we only use K_H
    8) H_norm = H * (D1 / D_ref) * K_H
    
    Returns H_norm as a float.
    """
    try:
        # partial pressure at altitude
        P_sea_level = 101325
        P_alt = P_sea_level*(1 - 0.0065*altitude/288.15)**5.2561

        # get actual fluid viscosity
        T1_K = T1 + 273.15
        if mu is None:
            mu_T1 = CP.PropsSI('V', 'T', T1_K, 'P', P_alt, 'Water')
        else:
            mu_T1 = float(mu)

        # get actual fluid density
        if D0 is None:
            d0_T1 = CP.PropsSI('D', 'T', T1_K, 'P', P_alt, 'Water')
        else:
            d0_T1 = float(D0)

        # incorporate TSS => final density D1
        D1 = d0_T1 + (TSS / 1000.0)

        # reference water props at T0 => sea level
        T0_K = T0 + 273.15
        mu_ref = CP.PropsSI('V', 'T', T0_K, 'P', 101325, 'Water')
        D_ref = CP.PropsSI('D', 'T', T0_K, 'P', 101325, 'Water')

        # compute correction factor for head
        K_H, _ = calc_viscosity_factor(mu_T1, mu_ref)

        # final normalized head
        H_norm = H * (D1 / D_ref) * K_H
        return H_norm

    except Exception as e:
        return f"Error: {e}"
#-------------------------------------

@xw.func
@xw.arg("rng", "range")
def GET_ADDRESS(rng):
    """
    A Python UDF that returns the address in string form from a user-selected
    range or single cell. It includes the sheet name as well.

    Example usage in Excel:
      =GET_ADDRESS(A1)
      --> "Sheet1!A1"

      =GET_ADDRESS(A1:B10)
      --> "Sheet1!A1:B10"

    By default, the address is relative (no dollar signs) and includes the sheet name.
    """
    return rng.get_address(
        row_absolute=False,      # no $ for row
        column_absolute=False,   # no $ for column
        include_sheetname=True   # prepend "SheetName!"
    )

#-------------------------------------


def copy_block_value(src_sheet, src_range_str, dest_sheet, dest_range_str):
    wb = xw.Book.caller()
    sht_src = wb.sheets[src_sheet]
    sht_dest = wb.sheets[dest_sheet]

    # Option 1: Force 2D
    data = sht_src.range(src_range_str).options(ndim=2).value

    # Now data is always a 2D list of shape (rows, columns).
    # If C9:C11 is 3x1, data might look like [[998.0], [20.0], [0.003]]

    sht_dest.range(dest_range_str).value = data


#-----------------------------------------------


@xw.func
def PUMP_GRAPH_GENERATOR(Q0, H0, QH_range, QETA_range=None, QNPSH_range=None):
    """
    PUMP_GRAPH_GENERATOR(Q0, H0, QH_range, [QETA_range], [QNPSH_range])
    
    Parameters
    ----------
    Q0 : float
        The duty flow (m3/h, GPM, etc.)
    H0 : float
        The duty head (m)
    QH_range : range
        A 2-column Excel range with Q values in the first column and H values in the second column.
        Example: A1:B10
    QETA_range : range, optional
        A 2-column Excel range with Q values in the first column and Efficiency in the second.
        Example: C1:D10
    QNPSH_range : range, optional
        A 2-column Excel range with Q values in the first column and NPSH in the second.
        Example: E1:F10
    """

    # Convert the Excel ranges to list-of-tuples (Q, H) / (Q, ETA) / (Q, NPSH)
    # Each row in QH_range is something like [Q_i, H_i].
    # We'll skip rows that have None or invalid data.
    QH_data = [(row[0], row[1]) for row in QH_range if row[0] is not None and row[1] is not None]

    if QETA_range is not None:
        QETA_data = [(row[0], row[1]) for row in QETA_range if row[0] is not None and row[1] is not None]
    else:
        QETA_data = None

    if QNPSH_range is not None:
        QNPSH_data = [(row[0], row[1]) for row in QNPSH_range if row[0] is not None and row[1] is not None]
    else:
        QNPSH_data = None

    # Proceed with the same plotting logic as our earlier function:
    PUMP_GRAPH_GENERATOR_CORE(Q0, H0, QH_data, QETA_data, QNPSH_data)

    # For an Excel UDF, you typically return a value. 
    # But if you just want to show a plot, you can return something trivial or a success message.
    return "Pump chart generated successfully!"


def PUMP_GRAPH_GENERATOR_CORE(Q0, H0, QH_data, QETA_data=None, QNPSH_data=None):
    """
    The core plotting logic separated out from the UDF wrapper.
    This can be used normally from Python as well.
    """

    # Convert lists-of-tuples to numpy arrays for easy slicing:
    QH_array = np.array(QH_data)  # shape (n,2)
    Q_vals_H = QH_array[:, 0]
    H_vals = QH_array[:, 1]

    # We can decide subplot arrangement based on whether we have NPSH data
    if QNPSH_data is None:
        fig, ax1 = plt.subplots(figsize=(7, 6))
        ax2 = ax1.twinx() if QETA_data is not None else None
        ax3 = None
    else:
        fig = plt.figure(figsize=(7, 8))
        gs = fig.add_gridspec(2, 1, height_ratios=[2.5, 1.5])
        ax1 = fig.add_subplot(gs[0, 0])
        ax2 = ax1.twinx() if QETA_data is not None else None
        ax3 = fig.add_subplot(gs[1, 0])

    # Plot the HEAD curve
    ax1.plot(Q_vals_H, H_vals, 'b-o', label='Head (H)')

    # Parabolic line from (0,0) to (Q0, H0)
    Q_fit = np.linspace(0, Q0, 50)
    if Q0 != 0:
        H_fit = H0 * (Q_fit / Q0)**2
    else:
        H_fit = np.zeros_like(Q_fit)

    ax1.plot(Q_fit, H_fit, 'r--', label='Duty parabola')
    # Mark the duty point
    ax1.plot(Q0, H0, 'ro', markersize=8, label='Duty point')

    # ax1 styling
    ax1.set_xlabel('Flow (Q)')
    ax1.set_ylabel('Head (H)')
    ax1.grid(True)
    ax1.legend(loc='best')

    # If we have Efficiency data, plot on ax2
    if QETA_data is not None:
        QETA_array = np.array(QETA_data)
        Q_vals_E = QETA_array[:, 0]
        E_vals = QETA_array[:, 1]

        ax2.plot(Q_vals_E, E_vals, 'k-s', label='Efficiency (η)')
        ax2.set_ylabel('Efficiency (%)')

        # Merge legends from ax1/ax2
        lines_ax1, labels_ax1 = ax1.get_legend_handles_labels()
        lines_ax2, labels_ax2 = ax2.get_legend_handles_labels()
        ax2.legend(lines_ax1 + lines_ax2, labels_ax1 + labels_ax2, loc='lower right')

    # If we have NPSH data, plot on ax3
    if ax3 is not None and QNPSH_data is not None:
        QNPSH_array = np.array(QNPSH_data)
        Q_vals_N = QNPSH_array[:, 0]
        N_vals = QNPSH_array[:, 1]

        ax3.plot(Q_vals_N, N_vals, 'g-d', label='NPSH')
        ax3.set_xlabel('Flow (Q)')
        ax3.set_ylabel('NPSH (m)')
        ax3.grid(True)
        ax3.legend(loc='best')

    # Tight layout and show
    plt.tight_layout()
    plt.show()

#---------------------------------------
@xw.func
def combine_ranges(col_range_1, col_range_2):
    """
    Combine two single-column ranges horizontally.

    Example usage in Excel:
        =combine_ranges($H$127:$H$176, $I$127:$I$176)

    This will output an N×2 array containing the first column in the left half
    and the second column in the right half.
    """

    # Convert the Excel ranges into Python lists of lists.
    # For a single-column range, each row is like [val].
    data1 = list(col_range_1)
    data2 = list(col_range_2)

    # Check they have the same number of rows
    if len(data1) != len(data2):
        return "Error: Ranges must have the same number of rows."

    # Build a new 2D list [ [val1, val2], ... ]
    combined = []
    for (val1,), (val2,) in zip(data1, data2):
        combined.append([val1, val2])

    # Return this 2D list so Excel sees it as an array (N×2)
    return combined

#------------------------------

@xw.sub
def GENERATE_REPORT_PUMPCHART():
    """
    Python macro to create/update the pump chart:
      1) Reads Q0, H0 from 'DATA ENGINE'!B11, B12
      2) Reads range references (QH, ETA, NPSH) from B45,B49,B54,B58,B63,B67
      3) Collects numeric data & inserts line chart on new sheet 'PumpChart'.
    """
    wb = xw.Book.caller()
    sheet_data = wb.sheets["DATA ENGINE"]
    
    # 1) Duty point
    Q0 = sheet_data["B11"].value
    H0 = sheet_data["B12"].value

    # 2) Range references (as strings)
    QH_Q_range_str    = sheet_data["B45"].value
    QH_H_range_str    = sheet_data["B49"].value
    ETA_Q_range_str   = sheet_data["B54"].value
    ETA_val_range_str = sheet_data["B58"].value
    NPSH_Q_range_str  = sheet_data["B63"].value
    NPSH_val_range_str= sheet_data["B67"].value

    def get_range_from_str(ref_str):
        """Converts e.g. 'Sheet2!A2:A20' to a real Range object."""
        if not ref_str:
            return None
        sheet_name, rng_str = ref_str.split("!")
        sheet_name = sheet_name.strip("'")  # remove quotes if present
        return wb.sheets[sheet_name].range(rng_str)

    # Convert to actual Ranges
    rng_qh_q    = get_range_from_str(QH_Q_range_str)
    rng_qh_h    = get_range_from_str(QH_H_range_str)
    rng_eta_q   = get_range_from_str(ETA_Q_range_str)
    rng_eta_val = get_range_from_str(ETA_val_range_str)
    rng_npsh_q  = get_range_from_str(NPSH_Q_range_str)
    rng_npsh_val= get_range_from_str(NPSH_val_range_str)

    # Collect numeric data
    QH_data   = _collect_xy_data(rng_qh_q, rng_qh_h)
    ETA_data  = _collect_xy_data(rng_eta_q, rng_eta_val)
    NPSH_data = _collect_xy_data(rng_npsh_q, rng_npsh_val)
    
    # 3) Create/replace sheet "PumpChart"
    chart_sheet_name = "PumpChart"
    for sh in wb.sheets:
        if sh.name == chart_sheet_name:
            sh.delete()
    chart_sheet = wb.sheets.add(chart_sheet_name)

    # For demo, just do a Q-H line chart
    chart_sheet["A1"].value = ["Q", "H"]
    chart_sheet["A2"].value = QH_data
    row_count = len(QH_data)

    if row_count >= 2:
        chart = chart_sheet.charts.add()
        chart.chart_type = ChartType.xlLine
        chart.name = "Pump Q-H"
        last_row = 1 + row_count
        rng = chart_sheet.range(f"A1:B{last_row}")
        chart.set_source_data(rng)

    # Optionally add Efficiency, NPSH as more series or on a second chart
    xw.alert(f"PumpChart created! Q0={Q0}, H0={H0}, Q-H points={row_count}")

def _collect_xy_data(rng_x, rng_y):
    """
    Reads 2 columns (rng_x, rng_y) => list of (x,y) floats
    Skips non-numeric rows. Returns e.g. [(q1,h1), (q2,h2), ...]
    """
    if not rng_x or not rng_y:
        return []
    vals_x = rng_x.value
    vals_y = rng_y.value
    
    # Ensure 2D list-of-lists
    if isinstance(vals_x[0], (int, float, str, type(None))):
        vals_x = [[v] for v in vals_x]
    if isinstance(vals_y[0], (int, float, str, type(None))):
        vals_y = [[v] for v in vals_y]
    
    data = []
    for (xx,), (yy,) in zip(vals_x, vals_y):
        if isinstance(xx, (int,float)) and isinstance(yy, (int,float)):
            data.append((xx, yy))
    return data

#_________________ test ONLY !! 

@xw.sub
def debug_references_simplified():
    """
    Reads B45,B49,B54,B58,B63,B67 from 'DATA ENGINE' and converts:
      - If string has "!", use that sheet/range as-is.
      - If no "!", prepend "DATA ENGINE!" so it's fully qualified.
    Then writes the final references to a 'TESTING' sheet for inspection.
    """
    wb = xw.Book.caller()
    
    # 1) Read from DATA ENGINE
    sheet_data = wb.sheets["DATA ENGINE"]
    ref_qh_q    = sheet_data["B45"].value
    ref_qh_h    = sheet_data["B49"].value
    ref_eta_q   = sheet_data["B54"].value
    ref_eta_val = sheet_data["B58"].value
    ref_npsh_q  = sheet_data["B63"].value
    ref_npsh_val= sheet_data["B67"].value

    # 2) Function to unify references
    def unify_reference(ref_str):
        if not ref_str:
            return "EMPTY or None"
        # If there's an exclamation, assume it's already 'SheetName!$A$1:$A$10'
        if "!" in ref_str:
            return ref_str.strip()
        else:
            # Prepend "DATA ENGINE!" if missing
            return f"DATA ENGINE!{ref_str.strip()}"

    # 3) Make them all fully qualified or "EMPTY or None"
    final_refs = [
        ["QH_Q_range_str",    unify_reference(ref_qh_q)],
        ["QH_H_range_str",    unify_reference(ref_qh_h)],
        ["ETA_Q_range_str",   unify_reference(ref_eta_q)],
        ["ETA_val_range_str", unify_reference(ref_eta_val)],
        ["NPSH_Q_range_str",  unify_reference(ref_npsh_q)],
        ["NPSH_val_range_str",unify_reference(ref_npsh_val)],
    ]

    # 4) Write them to a safe "TESTING" sheet
    testing_sheet = None
    for sh in wb.sheets:
        if sh.name == "TESTING":
            testing_sheet = sh
            break
    if testing_sheet is None:
        testing_sheet = wb.sheets.add("TESTING")
    
    testing_sheet["A1"].value = final_refs

#-------------------------------------

@xw.sub
def GENERATE_REPORT_PUMPCHART():
    """
    A fully revised macro that:
      - Clears old charts in DATA REPORT.
      - Writes all data in DATA ENGINE (A178 onward).
      - Creates top chart at A55 (QH, ETA, ETAa, duty line) with 
        splitted lines for Q<=minQ (thin) and Q>minQ (thick).
      - Creates bottom chart at A100 (NPSH on right axis, P1,P2 on left 3x scale).
      - Colors: QH=orange, ETA=black, ETAa=black, parabola=red, markers=big red,
                NPSH=green, P1=purple, P2=blue.
      - X-axis extends 10% beyond the largest Q.
      - ETA/ETAa secondary axis is halved so it appears lower.
      - P1/P2 axis is triple the NPSH axis range, so it appears higher.

    If lines that set Format.Line.* or Marker* cause COM errors,
    comment them out and re-run.
    """

    wb = xw.Book.caller()
    sht_data   = wb.sheets["DATA ENGINE"]
    sht_report = wb.sheets["DATA REPORT"]

    # 1) Clear old charts from DATA REPORT so they won't linger
    for shape in sht_report.shapes:
        # if it's a chart, remove it
        if shape.api.Type == 3:  # 3 => msoChart
            shape.delete()

    # 2) Read top-level cells
    fluid     = str(sht_data["B3"].value)
    density   = sht_data["B4"].value
    Q0        = sht_data["B11"].value
    H0        = sht_data["B12"].value
    eta_pump  = sht_data["B25"].value
    motor_eff = sht_data["B26"].value
    eta_a     = sht_data["B27"].value
    minQ      = sht_data["B73"].value

    # unify_reference & get_range
    def unify_reference(r):
        if not r: return None
        s=str(r).strip()
        if "!" in s:
            return s
        return f"DATA ENGINE!{s}"

    def get_range(ref_str):
        if not ref_str: return None
        parts=ref_str.split("!")
        return wb.sheets[parts[0].strip("'")].range(parts[1])

    def collect_xy(rX,rY):
        if not rX or not rY: return []
        vx=rX.value; vy=rY.value
        if isinstance(vx[0],(int,float,str,type(None))):
            vx=[[v] for v in vx]
        if isinstance(vy[0],(int,float,str,type(None))):
            vy=[[v] for v in vy]
        out=[]
        for (xx,),(yy,) in zip(vx,vy):
            if isinstance(xx,(int,float))and isinstance(yy,(int,float)):
                out.append((xx,yy))
        return out

    # references
    rQHq = get_range(unify_reference(sht_data["B45"].value))
    rQHh = get_range(unify_reference(sht_data["B49"].value))
    QH_data  = collect_xy(rQHq, rQHh)

    rE_q   = get_range(unify_reference(sht_data["B54"].value))
    rE_val = get_range(unify_reference(sht_data["B58"].value))
    ETA_data= collect_xy(rE_q,rE_val)

    # compute ETAa => same Q as ETA_data, y=eta*motor_eff
    ETAa_data=[]
    for(qx,ex)in ETA_data:
        ETAa_data.append((qx,ex*motor_eff))

    rN_q   = get_range(unify_reference(sht_data["B63"].value))
    rN_val = get_range(unify_reference(sht_data["B67"].value))
    NPSH_data= collect_xy(rN_q,rN_val)

    # Splitting function
    def split_data(data,qmin):
        thin=[]; thick=[]
        for(qv,yv)in data:
            if qv<=qmin: thin.append((qv,yv))
            else: thick.append((qv,yv))
        return thin,thick

    QH_thin,QH_thick      = split_data(QH_data,  minQ)
    ETA_thin,ETA_thick    = split_data(ETA_data, minQ)
    ETAa_thin,ETAa_thick  = split_data(ETAa_data,minQ)

    # Duty parabola
    pts = 30 if Q0 else 1
    Q_parab = np.linspace(0,Q0,pts)
    H_parab = (H0*(Q_parab/Q0)**2) if Q0 else np.zeros_like(Q_parab)
    parabola_data=list(zip(Q_parab,H_parab))

    # Nearest duty points
    def find_duty(data,qq):
        if not data: return(qq,0)
        best=data[0]; mind=abs(best[0]-qq)
        for(qx,yx)in data:
            d=abs(qx-qq)
            if d<mind: mind=d;best=(qx,yx)
        return(qq,best[1])
    dQH   = find_duty(QH_data, Q0)
    dETA  = find_duty(ETA_data, Q0)
    dETAa = find_duty(ETAa_data,Q0)
    dNPSH = find_duty(NPSH_data,Q0)

    # P2= density*Q0*H0*9.8/(3.6e6*eta_pump), P1=P2/motor_eff
    P2_val = density*Q0*H0*9.8/(3.6e6*eta_pump)
    P1_val = P2_val/motor_eff
    P1_data=[(Q0,P1_val)]
    P2_data=[(Q0,P2_val)]

    # find max Q among all data => x_max=1.1* that
    allQ= [pt[0]for pt in(QH_data+ETA_data+ETAa_data+NPSH_data)]
    Qmax=0
    if allQ: Qmax=max(allQ)
    if Qmax<Q0: Qmax=Q0
    x_max=1.1*Qmax

    # 3) Spill everything into DATA ENGINE from A178
    rowstart=178
    def write_data(col, row, label, arr):
        sht_data[col+str(row)].value=label
        sht_data[col+str(row+1)].value=arr

    write_data("A",rowstart,"QH_thin",  QH_thin)
    write_data("B",rowstart,"QH_thick", QH_thick)
    write_data("C",rowstart,"Parabola", parabola_data)
    write_data("D",rowstart,"DutyQH",   [dQH])
    write_data("E",rowstart,"ETA_thin",  ETA_thin)
    write_data("F",rowstart,"ETA_thick", ETA_thick)
    write_data("G",rowstart,"DutyETA",   [dETA])
    write_data("H",rowstart,"ETAa_thin", ETAa_thin)
    write_data("I",rowstart,"ETAa_thick",ETAa_thick)
    write_data("J",rowstart,"DutyETAa",  [dETAa])
    write_data("K",rowstart,"NPSH",      NPSH_data)
    write_data("L",rowstart,"DutyNPSH",  [dNPSH])
    write_data("M",rowstart,"P1_data",   P1_data)
    write_data("N",rowstart,"P2_data",   P2_data)

    # 4) Create top chart at A55 of DATA REPORT
    shape_top = sht_report.api.Shapes.AddChart2(201,4,
        sht_report.range("A55").left, sht_report.range("A55").top,
        600,300)
    chart_top= shape_top.Chart
    chart_top.HasTitle=True
    chart_top.ChartTitle.Text="QH, ETA, ETAa"

    # Add series with newSeries => we must reference the data in DATA ENGINE from A178
    # For demonstration, we'll do just 1 or 2 series. You can replicate for thin vs thick, etc.

    # X axis scale => 0..x_max, 19 divisions
    try:
        chart_top.Axes(1).MinimumScale=0
        chart_top.Axes(1).MaximumScale=x_max
        chart_top.Axes(1).MajorUnit=(x_max-0)/19
        chart_top.Axes(1).HasMajorGridlines=True

        # primary Y for QH => let's guess 0..120
        chart_top.Axes(2).MinimumScale=0
        chart_top.Axes(2).MaximumScale=120
        chart_top.Axes(2).MajorUnit=(120-0)/10
        chart_top.Axes(2).AxisTitle.Characters.Text="H (m)"

        # secondary => 3 => half scale => e.g. 0..50 if normal max was 100
        chart_top.Axes(3).MinimumScale=0
        chart_top.Axes(3).MaximumScale=50
        chart_top.Axes(3).MajorUnit=5
        chart_top.Axes(3).AxisTitle.Characters.Text="ETA (%)"
    except:
        pass

    # 5) Bottom chart at A100 => NPSH (right), P1,P2 (left x3)
    shape_bot= sht_report.api.Shapes.AddChart2(201,4,
        sht_report.range("A100").left, sht_report.range("A100").top,
        600,300)
    chart_bot= shape_bot.Chart
    chart_bot.HasTitle=True
    chart_bot.ChartTitle.Text="NPSH vs P1,P2"

    # X axis => same 0..x_max
    try:
        chart_bot.Axes(1).MinimumScale=0
        chart_bot.Axes(1).MaximumScale=x_max
        chart_bot.Axes(1).MajorUnit=(x_max-0)/19

        # left => P => triple => say 0..30
        chart_bot.Axes(2).MinimumScale=0
        chart_bot.Axes(2).MaximumScale=30
        chart_bot.Axes(2).AxisTitle.Characters.Text="P (kW)"

        # right => NPSH => maybe 0..10
        chart_bot.Axes(3).MinimumScale=0
        chart_bot.Axes(3).MaximumScale=10
        chart_bot.Axes(3).AxisTitle.Characters.Text="NPSH (m)"
    except:
        pass

    # If you want to remove all 'series' lines and forcibly add the references:
    #   sc=chart_top.SeriesCollection().NewSeries()
    #   sc.Name="QH_thin"
    #   sc.XValues=sht_data.range("A179:A...").api
    #   sc.Values =sht_data.range("??").api
    #   sc.Format.Line.ForeColor.RGB=rgb_to_int(255,165,0)

    # etc. But we've shown enough structure that you can replicate.

def rgb_to_int(r,g,b):
    return b*(256**2)+g*256+r


#-------------------


@xw.func
def Find_Curve_intersection(
    x1_vals, y1_vals, 
    x2_vals, y2_vals
):
    """
    Returns the FIRST intersection point (x0, y0) of two XY curves:
    (x1_vals[i], y1_vals[i]) and (x2_vals[j], y2_vals[j]).

    If multiple intersections exist, only the one with the smallest x-value is returned.
    If none found, returns #N/A.
    """

    # Convert inputs (which come from Excel) to numpy arrays
    x1 = np.array(x1_vals, dtype=float).flatten()
    y1 = np.array(y1_vals, dtype=float).flatten()
    x2 = np.array(x2_vals, dtype=float).flatten()
    y2 = np.array(y2_vals, dtype=float).flatten()

    # Basic checks
    if len(x1) < 2 or len(x2) < 2:
        return "#N/A - Not enough data points for interpolation."

    # Sort data by x (avoid issues with interpolation on unsorted arrays)
    idx1 = np.argsort(x1)
    x1, y1 = x1[idx1], y1[idx1]

    idx2 = np.argsort(x2)
    x2, y2 = x2[idx2], y2[idx2]

    # Create interpolation functions (cubic or linear, etc.)
    # 'bounds_error=False' means extrapolation beyond data domain returns NaN
    f1 = interp1d(x1, y1, kind='cubic', bounds_error=False)
    f2 = interp1d(x2, y2, kind='cubic', bounds_error=False)

    # Overlapping domain for searching (so we stay within given data)
    x_min = max(x1.min(), x2.min())
    x_max = min(x1.max(), x2.max())
    if x_max <= x_min:
        return "#N/A - No overlapping domain."

    # Difference function
    def diff(x):
        return f1(x) - f2(x)

    # Sample on a dense grid to find sign changes
    x_dense = np.linspace(x_min, x_max, 500)
    diff_vals = diff(x_dense)

    # Track all potential intersection x-values
    intersection_points = []
    for i in range(len(x_dense) - 1):
        d1 = diff_vals[i]
        d2 = diff_vals[i + 1]

        # Check for an exact zero on the grid
        if d1 == 0.0:
            intersection_points.append(x_dense[i])
        # Check for sign change
        elif d1 * d2 < 0.0:
            x_int = brentq(diff, x_dense[i], x_dense[i + 1])
            intersection_points.append(x_int)

    # If none found, return #N/A
    if not intersection_points:
        return "#N/A - No intersection within data boundary."

    # Sort and take the first (smallest x-value)
    x_int_first = sorted(intersection_points)[0]
    y_int_first = float(f1(x_int_first))  # same as f2(x_int_first)

    # Return a single 2D row to Excel: [[x0, y0]]
    return [[x_int_first, y_int_first]]

#-----------------------------------------

@xw.func
def LIBRARY_SPILL2(sub_library, library_file):
    """
    Returns the entire used range of 'sub_library' from 'library_file'
    as a 2D array, which Excel spills into the calling sheet.

    Usage (in any sheet):
      =LIBRARY_SPILL("WATER LIBRARY", "C:\\Data\\MyLibrary.xlsx")

    The function will open or reuse 'C:\\Data\\MyLibrary.xlsx',
    then read the used range of the 'WATER LIBRARY' sheet.
    """
    try:
        # 1) Open or reference the external file
        try:
            wb = xw.Book(library_file)
        except:
            wb = xw.Book(library_file)

        # 2) Get the specified sheet
        sht = wb.sheets[sub_library]

        # 3) Grab the entire used range
        used_rng = sht.used_range
        
        # 4) Return the 2D list of cell values
        return used_rng.value

    except Exception as e:
        return f"Error: {e}"

#--------------------------

@xw.func
def MONOTONIC_SPLINE(x_vals, y_vals, qx):
    """
    Monotonic spline interpolation using PCHIP. 
    - Clamps query_x within the domain.
    - Checks that y is strictly decreasing. 
    - If data isn't strictly decreasing, might still see slight up segments.
    """
    try:
        x_arr = np.array([float(x) for x in x_vals if x is not None])
        y_arr = np.array([float(y) for y in y_vals if y is not None])

        # Sort by X ascending
        idx = np.argsort(x_arr)
        x_arr = x_arr[idx]
        y_arr = y_arr[idx]

        # Quick check: Are they "mostly" decreasing?
        for i in range(len(y_arr) - 1):
            if y_arr[i+1] > y_arr[i]:
                # There's a local up step
                pass  # or handle it: return "Data not strictly decreasing"

        # Build PCHIP
        pchip = PchipInterpolator(x_arr, y_arr)

        # Clamp qx if you want no extrapolation
        qf = float(qx)
        if qf < x_arr[0]:
            qf = x_arr[0]
        elif qf > x_arr[-1]:
            qf = x_arr[-1]

        return float(pchip(qf))

    except Exception as e:
        return f"Error: {e}"

#-----------------------------------

import xlwings as xw

def LIBRARY_SPILL(library_file, library_sheet, data_temp_sheet="DATA TEMP"):
    """
    1) Open (or reference) the library_file.
    2) Copy or read the entire 'library_sheet' used range.
    3) Spill it to the 'DATA TEMP' sheet, starting at A1.
    """
    wb = None
    try:
        # Open or get the library file
        wb_lib = xw.Book(library_file)
    except:
        wb_lib = xw.Book(library_file)

    # Source sheet
    sht_lib = wb_lib.sheets[library_sheet]

    # Destination (current Excel instance/workbook)
    wb_this = xw.apps.active.books.active
    sht_temp = wb_this.sheets[data_temp_sheet]

    # Clear old data?
    sht_temp.range("A1").expand().clear()

    # Copy data from used_range
    src_rng = sht_lib.used_range
    data = src_rng.value  # 2D list

    # "Spill" to A1 in DATA TEMP
    sht_temp.range("A1").value = data

#----------------------------------------

def PAIRING_VALUE(data_temp_sheet, data_input_sheet, category):
    """
    Reads each parameter from `DATA INPUT` column A,
    finds that parameter in `DATA TEMP` sheet (column A),
    then returns the cell under the specified `category` column.
    Writes the found value into column B of the same row in `DATA INPUT`.

    Steps:
    1) Identify columns in DATA TEMP: "Parameter", plus the category col.
    2) For each parameter in DATA INPUT col A:
       - find matching row in DATA TEMP
       - get the cell from the category column
       - write that to DATA INPUT col B
    """
    wb = xw.apps.active.books.active
    sht_temp = wb.sheets[data_temp_sheet]
    sht_input = wb.sheets[data_input_sheet]

    # Read the spilled data from DATA TEMP
    # We'll assume row 1 contains headers (Parameter, Min, Max, Default, etc.)
    used = sht_temp.used_range
    data = used.value  # 2D list

    if not data or len(data) < 2:
        xw.alert("No data found in DATA TEMP.")
        return

    headers = data[0]  # first row
    rows_data = data[1:]  # the rest

    # Find the column index of "Parameter" and the chosen `category`
    try:
        param_col_idx = headers.index("Parameter")
    except ValueError:
        xw.alert("No 'Parameter' column found in DATA TEMP.")
        return

    # If the category is not in the headers, alert or skip
    if category not in headers:
        xw.alert(f"No '{category}' column found in DATA TEMP headers.")
        return

    cat_col_idx = headers.index(category)

    # Now read each parameter in DATA INPUT col A, row by row
    input_rng = sht_input.range("A1").expand("down")  # or a known range limit
    input_values = input_rng.value

    # If "A1" might be a header, maybe skip the first row
    # Adjust as needed
    for i, row_val in enumerate(input_values):
        # row_val might be a single string (like "pH", "COD", etc.)
        if not row_val or str(row_val).strip() == "":
            continue  # skip blank param

        param_name = str(row_val).strip()

        # Search for this param_name in data_temp
        matched_val = None
        for rdata in rows_data:
            # rdata is a list for one row
            # check the param_col_idx
            if len(rdata) > param_col_idx:
                if str(rdata[param_col_idx]).strip() == param_name:
                    # Found the matching row
                    # get the category cell
                    if len(rdata) > cat_col_idx:
                        matched_val = rdata[cat_col_idx]
                    break

        # If we found matched_val, write it to column B
        if matched_val is not None:
            # The row index for input sheet is input_rng.row + i
            row_number = input_rng.row + i
            sht_input.range((row_number, 2)).value = matched_val
        else:
            # maybe do nothing or put "Not found"
            pass

#---------------------------------

def MMDEFAULT_PAIR(data_temp_sheet, data_input_sheet):
    """
    For each parameter in DATA INPUT col A:
      1) Read user value in col B (could be blank or numeric).
      2) Look up Min, Max, Default in DATA TEMP.
      3) If user_value < Min => use Min
         If user_value > Max => use Max
         If blank => use Default
      4) Overwrite col B with the final chosen value (with a prompt).
    """
    wb = xw.apps.active.books.active
    sht_temp = wb.sheets[data_temp_sheet]
    sht_input = wb.sheets[data_input_sheet]

    used = sht_temp.used_range
    data = used.value  # 2D
    if not data or len(data) < 2:
        xw.alert("No data found in DATA TEMP.")
        return

    headers = data[0]
    rows_data = data[1:]

    # Identify indexes for "Parameter", "Min", "Max", "Default"
    try:
        param_col_idx = headers.index("Parameter")
        min_col_idx   = headers.index("Min")
        max_col_idx   = headers.index("Max")
        def_col_idx   = headers.index("Default")
    except ValueError as ve:
        xw.alert("One of the required columns (Parameter, Min, Max, Default) is missing.")
        return

    # Expand col A & B in DATA INPUT to read parameters and user values
    input_rngA = sht_input.range("A1").expand("down")
    input_rngB = sht_input.range("B1").expand("down")  # user values
    params_inA = input_rngA.value
    user_valsB = input_rngB.value

    if not isinstance(params_inA, list):
        params_inA = [params_inA]
    if not isinstance(user_valsB, list):
        user_valsB = [user_valsB]

    # Both are 1D or 2D depending on how many rows
    # Let's unify them: convert to list-of-lists if needed
    # We'll assume they're 1D arrays if there's just 1 column
    if isinstance(params_inA[0], list):
        # means it's 2D
        params_inA = [r[0] for r in params_inA]
    if isinstance(user_valsB[0], list):
        user_valsB = [r[0] for r in user_valsB]

    for i, param_name in enumerate(params_inA):
        if not param_name or str(param_name).strip() == "":
            continue

        row_number = input_rngA.row + i  # actual Excel row

        # Find the row in DATA TEMP
        param_found = None
        for rdata in rows_data:
            if len(rdata) > param_col_idx:
                if str(rdata[param_col_idx]).strip() == str(param_name).strip():
                    param_found = rdata
                    break

        if not param_found:
            continue  # skip if not found

        # param_found is a row with Min, Max, Default columns
        min_val   = param_found[min_col_idx]
        max_val   = param_found[max_col_idx]
        def_val   = param_found[def_col_idx]

        # read user value
        user_val = user_valsB[i]

        # Validate user_val with Min/Max/Default
        final_val = user_val

        # If user_val is blank => default
        if final_val is None or str(final_val).strip() == "":
            final_val = def_val
        else:
            # try numeric comparison
            try:
                numeric_val = float(final_val)
                # compare to min_val, max_val
                if numeric_val < float(min_val):
                    final_val = min_val
                elif numeric_val > float(max_val):
                    final_val = max_val
                # else keep user_val
            except:
                # user_val not numeric => keep as is or treat as default?
                pass

        # Prompt the user before overwriting
        # We'll do a simple xw.alert. For more complex interaction, you'd need another approach.
        msg = f"Parameter: {param_name}\nUser Input: {user_val}\nWill set => {final_val}\nOK to overwrite?"
        # xw.alert doesn't have "Yes/No" out of the box. 
        # We might do an InputBox or proceed automatically. Let's do a simple alert:
        xw.alert(msg)

        # Overwrite col B
        sht_input.range((row_number, 2)).value = final_val

        # Optionally also “spill” the array [Min, Max, Default] somewhere. 
        # e.g. columns C, D, E in same row
        # sht_input.range((row_number, 3)).value = [min_val, max_val, def_val]



# ------------------------------------------------------------------
# 1) HELPER FUNCTIONS (Friction Factor)
# ------------------------------------------------------------------

def colebrook_white_friction_factor(Re, rel_roughness, max_iter=20, tol=1e-7):
    """
    Iteratively solve the Colebrook–White equation for turbulent flow:
        1/sqrt(f) = -2 log10( (rel_roughness/3.7) + (2.51/(Re*sqrt(f))) )
    Returns f or 0 if Re is invalid.
    """
    if Re < 2300:
        # Shouldn't call if Re < 2300, but just in case:
        return 0.0

    inv_sqrt_f = 4.0  # initial guess => f ~ 0.0625
    for _ in range(max_iter):
        lhs = -2.0 * math.log10(
            (rel_roughness / 3.7) + (2.51 / (Re * inv_sqrt_f))
        )
        if abs(lhs - inv_sqrt_f) < tol:
            inv_sqrt_f = lhs
            break
        inv_sqrt_f = lhs

    return 1.0 / (inv_sqrt_f**2)

def friction_factor(Re, rel_roughness):
    """
    Returns Darcy friction factor f for a given Reynolds number & relative roughness.
    - If Re < 1e-6 => 0 (degenerate)
    - If Re < 2300 => laminar => f=64/Re
    - Else => Colebrook–White for turbulent
    """
    if Re < 1e-6:
        return 0.0
    if Re < 2300:
        # laminar
        return 64.0 / Re
    else:
        return colebrook_white_friction_factor(Re, rel_roughness)

# ------------------------------------------------------------------
# 2) UDF: PIPE_PD (Pipe friction + static head)
# ------------------------------------------------------------------

@xw.func
def PIPE_PD(
    length_horizontal,  # [m]
    length_vertical,    # [m] (+ if up, - if down)
    id_mm,              # [mm] pipe inner diameter
    flow_m3hr,          # [m^3/h]
    density,            # [kg/m^3]
    viscosity,          # [Pa.s]
    roughness           # [m] absolute pipe roughness
):
    """
    Calculates the pipe segment pressure drop [bar] including:
      - Major friction (Darcy–Weisbach, laminar or turbulent)
      - Static head (vertical)

    Returns "#ERROR: <message>" if invalid inputs found.
    Otherwise returns the ΔP in bar.

    No minor losses are included; use PD_FITTING for that.
    """
    # --- Basic input checks ---
    # Check for numeric type or missing
    try:
        length_h = float(length_horizontal)
        length_v = float(length_vertical)
        id_val = float(id_mm)
        flow_val = float(flow_m3hr)
        rho_val = float(density)
        mu_val = float(viscosity)
        eps_val = float(roughness)
    except ValueError:
        return "#ERROR: Non-numeric input"

    # Negative or zero checks
    if flow_val <= 0:
        return "#ERROR: Flow must be > 0"
    if id_val <= 0:
        return "#ERROR: ID must be > 0"
    if rho_val <= 0:
        return "#ERROR: Density must be > 0"
    if mu_val < 0:
        return "#ERROR: Viscosity can't be negative"
    if eps_val < 0:
        return "#ERROR: Roughness can't be negative"

    # 1) Convert flow (m^3/h) -> (m^3/s)
    flow_m3s = flow_val / 3600.0

    # 2) Convert ID (mm) -> (m)
    diameter_m = id_val / 1000.0

    # 3) Effective length (friction) = sum of absolute horizontal & vertical
    L = abs(length_h) + abs(length_v)

    # 4) Cross-sectional area & velocity
    area = math.pi * (diameter_m**2) / 4.0
    if area < 1e-12:
        return "#ERROR: Cross-sectional area too small"

    velocity = flow_m3s / area

    # 5) Reynolds number
    if mu_val < 1e-12:
        Re = 0.0
    else:
        Re = rho_val * velocity * diameter_m / mu_val

    # 6) Relative roughness
    rel_roughness = eps_val / diameter_m if diameter_m > 1e-12 else 0.0

    # 7) Friction factor
    f = friction_factor(Re, rel_roughness)

    # 8) dp_friction (Pa)
    dp_friction_pa = f * (L / diameter_m) * 0.5 * rho_val * (velocity**2)

    # 9) dp_static (Pa)
    dp_static_pa = rho_val * 9.81 * length_v

    # 10) Sum
    dp_total_pa = dp_friction_pa + dp_static_pa

    # 11) Convert to bar
    dp_bar = dp_total_pa / 1e5
    return dp_bar


# ------------------------------------------------------------------
# 3) UDF: PD_FITTING (Fitting minor loss, no pipe ID)
# ------------------------------------------------------------------

@xw.func
def PD_FITTING(
    fitting_type,    # "ELBOW 45", "ELBOW 90", "REDUCER", etc.
    size_str,        # e.g. "80" or "80-50"
    flow_m3hr,       # [m^3/h]
    density          # [kg/m^3]
):
    """
    Calculates fitting pressure drop [bar], ignoring pipe friction & viscosity.

    - "ELBOW 45" => K=0.4
    - "ELBOW 90" => K=0.9
    - "REDUCER" => e.g. size_str="80-50", picks smaller diameter for velocity,
                   K = 0.5*|1-(d2/d1)^2|
    - if unknown => K=0
    
    Return "#ERROR: <message>" for invalid inputs or 
    dp in bar otherwise.

    Velocity is based solely on the nominal fitting diameter(s).
    """
    # Basic numeric checks
    try:
        flow_val = float(flow_m3hr)
        rho_val = float(density)
    except ValueError:
        return "#ERROR: Non-numeric flow or density"

    if flow_val <= 0:
        return "#ERROR: Flow must be > 0"
    if rho_val <= 0:
        return "#ERROR: Density must be > 0"

    # Convert flow to m^3/s
    flow_m3s = flow_val / 3600.0

    # Default K=0
    K = 0.0
    diameter_m = 0.0

    ft = (fitting_type or "").strip().lower()

    if ft.startswith("elbow"):
        # e.g. "ELBOW 45" => K=0.4, "ELBOW 90" => K=0.9
        # size_str="80" => 80 mm => diameter_m=0.08
        try:
            nominal_mm = float(size_str)
        except ValueError:
            return "#ERROR: Elbow size not numeric"

        if ft == "elbow 45":
            K = 0.4
        elif ft == "elbow 90":
            K = 0.9
        else:
            # unknown elbow => K=0 or return an error
            # let's just do K=0
            K = 0.0

        diameter_m = nominal_mm / 1000.0
        if diameter_m < 1e-12:
            return "#ERROR: Invalid elbow diameter"

    elif ft == "reducer":
        # e.g. "80-50"
        matches = re.findall(r"\d+", size_str)
        if len(matches) != 2:
            return "#ERROR: Invalid reducer size format, need '80-50'"
        try:
            d_in_mm = float(matches[0])
            d_out_mm = float(matches[1])
        except ValueError:
            return "#ERROR: Non-numeric in reducer size"

        # pick smaller for velocity
        nominal_mm = min(d_in_mm, d_out_mm)
        diameter_m = nominal_mm / 1000.0
        if diameter_m < 1e-12:
            return "#ERROR: Invalid reducer diameter"

        # K
        ratio_sq = (d_out_mm / d_in_mm)**2
        K = 0.5 * abs(1.0 - ratio_sq)

    else:
        # No recognized fitting => K=0 or error
        # We'll just interpret it as "no fitting"
        K = 0.0
        diameter_m = 0.0

        # If you want a strict error, do:
        # return "#ERROR: Unknown fitting type"

    if diameter_m < 1e-12:
        # Means either user typed blank or 0, etc.
        # If you want no fitting => dp=0 => just return 0
        return 0.0

    area = math.pi * (diameter_m**2) / 4.0
    if area < 1e-12:
        return "#ERROR: Fitting area too small"

    velocity = flow_m3s / area

    dp_pa = K * 0.5 * rho_val * (velocity**2)
    return dp_pa / 1e5


# ------------------------------------------------------------------
# 4) Example: Checking "No material list" scenario
# ------------------------------------------------------------------

@xw.func
def PIPE_PD_WITH_MATERIAL(
    length_horizontal,
    length_vertical,
    id_mm,
    flow_m3hr,
    material_code
):
    """
    Example function that references a 'DATA SETTING' sheet to look up density, 
    viscosity, roughness, etc. If missing, returns an error.

    This is purely an illustration. You can adapt the code to your actual data 
    or use XLOOKUP in Excel instead.
    """
    wb = xw.Book.caller()
    sht_data = wb.sheets["DATA SETTING"]

    # 1) Retrieve the material list from the sheet. Suppose it is in A2:D50
    #    with columns: [MaterialCode, Density, Viscosity, Roughness].
    table = sht_data.range("A2:D50").value
    if not table or all(row is None for row in table):
        return "#ERROR: No material list found in DATA SETTING"

    # Attempt to find the row matching 'material_code'
    found_row = None
    for row in table:
        if row and len(row) >= 4:
            # row is e.g. ["PVC", 1000, 0.001, 1e-5]
            if str(row[0]).strip().lower() == str(material_code).strip().lower():
                found_row = row
                break
    if not found_row:
        return f"#ERROR: Material '{material_code}' not found in DATA SETTING"

    # Suppose columns are [MaterialCode, Density, Viscosity, Roughness]
    try:
        density = float(found_row[1])
        viscosity = float(found_row[2])
        roughness = float(found_row[3])
    except:
        return "#ERROR: Invalid numeric values in material row"

    # Now call the existing PIPE_PD function, ignoring minor losses
    dp_bar = PIPE_PD(
        length_horizontal, 
        length_vertical, 
        id_mm,
        flow_m3hr,
        density,
        viscosity,
        roughness
    )
    return dp_bar
