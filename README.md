# EPU_Screening_Visualization
A web app using Flask that can run on a support PC and allow users to visualize screening and collection sessions running in a single particle session on EPU. Compatible with Tundra electron microscopes and should be mostly compatible with Krios and Talos microscopes.

# Getting Started
1. This app requires that Windows Subsystem for Linux (WSL) and miniconda are downloaded on the support PC
2. Generate a conda environment appropriate for the screening visualization
   conda create --name screening_vis python=3.9 reportlab Flask pandas numpy Pillow
3. conda activate screening_vis
4. The code assumes that your screening and collection data are mounted at /mnt/z on the support PC. If this is not the case, modify the two locations in app.py that specify /mnt/z as the base root.
5. The code assumes that you have a pixel size table named pixelsizes.txt located at the location of the base root. There is an example file in the code here that you can modify.

# Running the Script
cd /path/to/screening_vis
conda activate screening_vis
python app.py
Copy and paste the http://XXX.X.X.X:XXXX into the web browser on the support PC and it should allow you to interactively browse screening and data collection sessions
