# Baseline Parameter Automation
Automation of baseline parameters configuration for RAN nodes deployment and integration. This application allows to automate RAN parameter definition and configuration, saving time and decreasing the probability of making errors. Python and OpenPyXL are used.

# Requirements
- Input file must be in "data" folder and has to be named "input.xlsx".
- Managed Objects (MO) GNBCUCPFunction, CUEUtranCellFDDLTE and NRCellDU must contain the necessary cell information of the target nodes.

# How to use
An executable file is provided (ZTE Baseline Tool.exe). Besides, the application can be build using PyInstaller package. All or individual steps can be selected and the output can be found in "data" folder.
The available steps are:
- NRFREQ: Parameters to define the frequency band and associate it with the new cells, as well as the allocation of time-frequency resources.
- EXTQCI: Quality of service for different traffic classes by assigning QoS Class Identifier to Data Radio Bearers.
- 4G-4G Cosites: Configuration of neighboring 4G-4G cosite cells for the correct execution of intra-eNodeB handover and carrier aggregation.
- 5G-5G Cosites: Configuration of neighboring 5G-5G cosite cells for the correct execution of intra-gNodeB handover.
