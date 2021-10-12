# VBASubnetFinderChecker
When you are in a datacenter and you have a long list of existing subnets and VLANs, you may need to find out which subnet your IP belongs to.
Here, we use an Excel and a simple VBA Code to help you solving that problem.
The Excel file is used as "workbench": you paste there all the IPs to recognize and all the subnets you have. 

The VBA code does the following: for every unknown IP in the list, the vba engine compares the single IPv4 Octects of the "unknown" IP with each "existing subnet" whether the IP octects values are in between the first IP of the subnet and the last IP of the subnet.

The result is the list of "matched" subnets.
