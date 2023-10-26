# VBASubnetFinderChecker
When you are in a datacenter and you have a long list of existing subnets and VLANs, you may need to find out which subnet your IP belongs to.
Here, we use an Excel file and a simple VBA Code to help you solving that problem.
The Excel file is used as "workbench": you paste there all the IPs you need to classify and in another column, all the subnets you have. 

The VBA code does the following: for every IP in the first column, the vba engine compares the single IPv4 Octects of the "unknown" IP with each "existing subnet": it checks whether the IP octects values are in between the first IP and the last IP of each subnet.

The result is the list of "matched" subnets so the unknown IP is now correctly classified.
