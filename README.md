# Hysteresis-Calculator
Program that performs the mathematical analysis of hysteresis behavior detected in voltage gating of large beta-barrel transmembrane ion channels

## Project Description
A desktop application developed for the NICHD/DIR Section on Molecular Transport to perform mathematical analysis on ramp data, 
specifically for analyzing hysteresis behavior in the voltage gating of large beta-barrel transmembrane ion channels. 
Prior to the development of this application, the lab had to rely on manual computations, which were time-consuming and labor-intensive.

## Installation
Navigate to the the Application Setup folder in the main repository and download th eexecutable to gain access to the desktop application

## Usage
Data Formating:
1. The input file must be a '.xlsx' file
2. The file should be formatted as such:
    Column 1: Time
    Column 2: Current
    Column 3: Voltage
3. When formatting include rows 0 to max_value including 0 and discluding the max_value
    ex. data on inter 0 to 80,000
    Notice how the 80,000th value is discluded as this is counted as the first value in the next ramp
    
    0	-0.00005607	-0.004029863
    0.25	-0.000189189	-0.003831981
    0.375	-0.000713006	-0.003298745
    0.625	-0.001476001	-0.002481842
    0.875	-0.002408739	-0.001484072
    ...
    ...
    ...
    79998.25	-0.017886072	0.006433671
    79998.375	-0.01835297	0.007634072
    79998.625	-0.018580325	0.008664874
    79998.875	-0.01858563	0.009583855
    79999	-0.018413454	0.010443585

Landing Page: 
1. Add the ramp data under "Data(.xlsx)"
2. Add the ramp number under "Number of Ramps"
3. Click "Run Analysis"
Results Page:
1. Click "Save Single Ramp Data" to save the average of every ramp in your desired file location
2. Click "Save Overall Ramp Data" to save partitioned ramp data and average ramp data in your desired file location
3. Click "Return to Main Page" to return to the Main Page and continue using the application

## Contact Information
For any help, information, bug reports, or feature requests, please feel free to contact me at tmertz@usc.edu
I am happy to assist and address any concerns you may have.
