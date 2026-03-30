from __future__ import division

import math
import os
import xlwings as xw
import numpy as np
import pandas as pd
from matplotlib import pyplot as plt
from numpy.linalg import inv, pinv
from scipy.optimize import minimize
from colorama import Fore, Style, init
init(autoreset=True)

# Find Workbook
current_directory = os.path.dirname(os.path.abspath(__file__))
workbook_filename = "PORTFOLIO ALLOCATION MODEL.xlsm"
workbook_path = os.path.join(current_directory, workbook_filename)
wb = xw.Book(workbook_path)
print('\n' "Current Directory:" '\n', current_directory)
print("Workbook Path:" '\n', workbook_path, '\n')

# Define n
ws = wb.sheets["Security Entry"]
n = ws.range("B37").value
n = int(n)  # Convert to integer if needed

# Print to verify the value of n
print(Fore.LIGHTYELLOW_EX + "NUMBER OF ASSETS (n):" '\n', n, '\n')

# Define OMEGA
ws_omega = wb.sheets["Matrixes"]
omega_range = ws_omega.range("VarCovarMatrix")  # OMEGA matrix data
omega = np.matrix(omega_range.value)
if n < 30:
    omega = omega[:n, :n]
display_omega = pd.DataFrame(omega)
display_omega.index = display_omega.index + 1
display_omega.columns = display_omega.columns + 1
print(Fore.LIGHTYELLOW_EX + "OMEGA MATRIX:\n", display_omega)

# Define Empty W Variable
W = [1/n] * n
print(Fore.LIGHTCYAN_EX + "\nInitialised Ws  (1/n):")
print(W)
print(Fore.BLUE + "Sum of Initial Ws: ", np.sum(W))

# risk budgeting optimization
def calculate_portfolio_var(W, omega):
    # function that calculates portfolio risk
    W = np.matrix(W)
    return math.sqrt((W*omega*W.T)[0,0])
sigma = calculate_portfolio_var(W, omega)
print(Fore.LIGHTCYAN_EX + "\nInitial Portfolio Volatility: ", sigma)


def calculate_risk_contribution(W, omega):
    # function that calculates asset contribution to total risk
    W = np.matrix(W)
    sigma = np.sqrt(calculate_portfolio_var(W,omega))
    if sigma == 0:
        return np.zeros_like(W)
    # Marginal Risk Contribution
    MRC = (omega*W.T)/sigma
    # Risk Contribution
    RC = np.multiply(MRC,(W.T/sigma))
    return RC/sigma**2
rc_values = calculate_risk_contribution(W, omega)
print(Fore.LIGHTCYAN_EX + "\nInitial Risk Contributions:")
for i, rc in enumerate(rc_values):
    print(f"RC {i+1}: {rc}")
print(Fore.BLUE + "Sum of Initial RCs: ", np.sum(rc_values))
print()

def risk_budget_objective(x,pars):
    # calculate portfolio risk
    omega = pars[0] # covariance table
    x_t = pars[1] # risk target in percent of portfolio risk
    sig_p = np.sqrt(calculate_portfolio_var(x,omega)) # portfolio sigma
    risk_target = np.asmatrix(np.multiply(sig_p,x_t))
    asset_RC = calculate_risk_contribution(x,omega)
    J = sum(np.square(asset_RC-risk_target.T))[0,0] * 1000 # sum of squared error
    return J

def total_weight_constraint(x):
    return np.sum(x)-1.0

def long_only_constraint(x):
    return x

x_t = [1/n] * n # your risk budget percent of total portfolio risk (equal risk)
cons = ({'type': 'eq', 'fun': total_weight_constraint},
{'type': 'ineq', 'fun': long_only_constraint})
res = minimize(risk_budget_objective,W , args=[omega,x_t], method='SLSQP',constraints=cons, options={'disp': True, 'ftol': 1e-12})
W_rb = np.asmatrix(res.x)
print(Fore.LIGHTGREEN_EX + "\nOPTIMISED Ws:")
print(W_rb)
print(Fore.GREEN + "Sum of Optimised Ws: ", np.sum(W_rb))

erc_sigma = calculate_portfolio_var(W_rb, omega)
print(Fore.LIGHTGREEN_EX + "\nOPTIMISED Portfolio Volatility: ", erc_sigma)

optimised_rc_values = calculate_risk_contribution(W_rb, omega)
print(Fore.LIGHTGREEN_EX + "\nOPTIMISED Risk Contributions:")
for i, rc in enumerate(optimised_rc_values):
    print(f"RC {i+1}: {rc}")
print(Fore.GREEN + "Sum of Optimised RCs: ", np.sum(optimised_rc_values))

# Plug Optimized ERC Ws into Excel Doc
W_rb = np.array(W_rb)
ws_portfolio = wb.sheets["Portfolio"]
ws_portfolio.range("ERC_Ws").options(transpose=True).value = W_rb