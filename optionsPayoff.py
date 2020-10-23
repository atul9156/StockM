import sys
import pandas as pd
import numpy as np
import plotly.graph_objects as go

def payoff(path, sheet_name = "OPTIONS", onlyPayoff = False):
	"""
	Function to dispaly the payoff diagram of an option stategy
	parameters:
	- path: Path of the excel file which has the type, strike price and premium for the options involved
	- sheet_name: Name of the worksheet in the excel workbook which contains Options info
	- onlyPayoff: If set to True, only the cumulative payoff of the Options Strategy will be displayed
	"""

	df = pd.read_excel(path, sheet_name = sheet_name)
	ms = df['strike'].mean()
	prange = np.linspace(df["strike"].min() - 0.1 * ms, df["strike"].max() + 0.1 * ms, 100)
	cumPayoff = np.array([0 for x in prange], dtype = "float64")
	line = 0
	# graph object or matplotlib figure
	fig = go.Figure().update_layout(
			title = "Option Payoff",
			xaxis_title = "Spot Price",
			yaxis_title = "Profit/Loss"
			)
	for item in df.itertuples():
		typ, s, p, lots, pos = item.type, item.strike, item.premium, item.lot_size, item.position
		line += 1
		if typ.upper() == "CE" and pos.upper() == "B":
			value = np.array([max(-p, x - (s + p)) * lots for x in prange])
		elif typ.upper() == "CE" and pos.upper() == "S":
			value = np.array([min(p, (s + p) -x) * lots for x in prange])
		elif typ.upper() == "PE" and pos.upper() == "B":
			value = np.array([max(-p, (s - p) - x) * lots for x in prange])
		elif typ.upper() == "PE" and pos.upper() == "S":
			value = np.array([min(p, x - (s - p)) * lots for x in prange])
		else:
			if typ.upper() not in ["PE", "CE"]:
				print("Option type not recognized, ERROR at line {}".format(line))
				print("Option type entered:", typ)
				print("Valid values must be from the following tuple: (PE, CE)")
				sys.exit("Option type undefined")
			else:
				print("Position type not recognized, ERROR at line {}".format(line))
				print("Position type entered:", pos)
				print("Valid values must be from the following tuple: (B, S)")
				sys.exit("Position type undefined")

		cumPayoff = np.add(cumPayoff, value)
		key = "{0}_{1}".format(typ, s)

		if onlyPayoff == False:
			fig.add_trace(
				go.Scatter(
					x = prange,
					y = value,
					name = key)
			)

	if not (onlyPayoff == False and line == 1):
		fig.add_trace(
			go.Scatter(
				x = prange,
				y = cumPayoff,
				name = "Cumulative PayOff")
			)
	fig.show()


path = input("Enter the path of you Excel file:")
sheet_name = input("Enter name of sheet which contains Options info:")
payoffOnly = input("Do you only want to see the Cumulative payoff of your Options Strategy? [y/n]:")

if payoffOnly == "y":
	payoff(path, sheet_name, True)
else:
	payoff(path, sheet_name)
