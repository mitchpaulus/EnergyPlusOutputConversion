# EnergyPlusOutputConversion
Executable that converts the .csv output from EnergyPlus into an .xlsx file and performs common unit conversions along with some formatting.

## What does it all do?

The program will open a file prompt, in which it is expected that you select an EnergyPlus output .csv file. It then converts that file to an Excel format, .xlsx. 

It then

1. Word wraps the first row.
2. Converts any "Electricity" columns to kWh, and formats with 0 decimal places.
3. Converts any "DistrictCooling" or "DistrictHeating" columns to MMBTU, and formats with 0 decimal places.
4. Converts any "Temperature" columns to F, and formats with 1 decimal place.
5. Adjusts the width of all the columns. 
6. Saves the file with -convert.xlsx as the new ending in the same folder the original file was in.

## How can I use it?

After you download the zip file, you can run the program directly without installing any program to your computer by running the executable in \EnergyPlusConverter\EnergyPlusConverter\bin\Debug\EnergyPlusConverter.exe. 

The other option is to run the setup.exe file, located at \EnergyPlusConverter\EnergyPlusConverter\publish\setup.exe, which will then install the program on your computer along with a shortcut. 
