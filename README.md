# powershell
Collection of Powershell Scripts/Modules

Write-Log.psm1 - Module used in almost every script for logging messages, includes functions to start logging, write log messages and stop logging

MDT_Drivers - A set of scripts and files that I've used to update the drivers in MDT Deployment shares. These are set up to use mdt in the "Total Control" driver method. There are a number of pieces there so here's an overview of them
- Manifests folder: the driver manifests generated by the scripts, simply a listing of all the models and the drivers downloaded for each one
- Models folder: contains the csv files for the models you want to update: currently one for HP and one for Dell.
- Scripts folder: where the main scripts are stored, 4 of them currently
  - Clear_Old_Driver_Folders - Checks through the driver folders and clears out ones that haven't been updated in awhile. 
  - Run_MDT_Driver_Updates - The main script that runs the other ones, you set the deployment share to target here and it's used                              as a parameter for the other scripts
  - Update_Dell_Drivers_MDT - Runs the driver update for the models in the Dell_Driver_Models csv file
  - Update_HP_Drivers_MDT - Runs the driver update for the models in the HP_Driver_Models csv file
