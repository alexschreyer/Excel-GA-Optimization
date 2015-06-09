Program:   GA Optimization for EXCEL
Version:   1.2
Author:   Alexander C. Schreyer
Copyright:   © 2005, Alexander C. Schreyer
Website:   http://www.alexschreyer.de/projects/xloptim/
E-mail:   mail@alexschreyer.de

DESCRIPTION:
--------------------
This program performs a Genetic Algorithm (GA) optimization of a Microsoft Excel spreadsheet. The user can specify the type of optimization (maximization, minimization, target value) and modify parameters related to the GA. Due to the nature of this robust optimization routine, the calculation model may be nonlinear and/or discontinuous. The model may contain five design variables and five constraints (of the "<=", ">=" or "=" type). Design variables can take real or integer values within specified lower and upper bounds.

For the time being, I omitted a help-file since the UI should be self-explanatory. If there is a demand for one, I will publish it as well.

VERSION:
--------------------
1.2 (August 2005) - Build 1018

LICENSE:
--------------------
By installing and using the software (“software” means all programs and files associated with this program), you (“you” means the installer and all users) agree to abide by the terms outlined in the LICENSE.TXT file.

KNOWN BUGS / LIMITATIONS (please send comments to mail@alexschreyer.de):
--------------------
- If you have less than five design variables or constraints in your model, then assign "fake cells" to those. Put a number (e.g. 10) into the cells and set the variable's limits (or a constraint's target value) to that number (e.g. LL = 10 and UL = 10).
- If a "divide by zero" error may be encountered in Excel when a design variable becomes zero, then specify a close-enough limit (e.g. 0.001) to prevent this from happening.
- The constraint penalty is currently fixed at 10^12.

REVISION HISTORY:
--------------------
See the website for details.

====================
A. Schreyer, August 2005
