# EPU_Screening_Visualization
A web app using Flask that can run on a support PC and allow users to visualize screening and collection sessions running in a single particle session on EPU. Compatible with Tundra electron microscopes and should be mostly compatible with Krios and Talos microscopes.

# Getting Started
1. This app requires that Windows Subsystem for Linux (WSL) and miniconda are downloaded on the support PC
2. Generate a conda environment appropriate for the screening visualization
   ```bash
   conda create --name screening_vis python=3.9 reportlab Flask pandas numpy Pillow
   ```
5. Activate the conda environment
   ```bash
   conda activate screening_vis
   ```
6. Modify app.py as needed
      - The code assumes that your screening and collection data are mounted at /mnt/z on the support PC. If this is not the case, modify the two locations in app.py that specify /mnt/z as the base root.
7. Modify epu/epustats.py as needed
      - First, change
      ```python
      MICROSCOPE_INFO = {
         "TUNDRA-XXX": ("DFCI Tundra", 1.6),
         "TITANXXX": ("HMS Krios2", 2.7),
         "TITANXXX": ("HMS Krios1", 2.7),
      }
      ```
      to contain your serial number in the XXX spot and the correct spherical aberration in mm for your microscope. Note that the Tundra has a hyphen after it whereas the Titan does not. 
      - Next, if you will be using this script with a Tundra, modify the instances of
      ```python
      if instrument_model == "TUNDRA-XXX":
      ```
      to contain your serial number in the XXX spot. 
      - Lastly, depending on your image format, you may have to modify this snippet to change the extension:
      ```python
      if instrument_model == "TUNDRA-XXX":
         fractions_ext = "mrc"
         pattern = "*Fractions.mrc"
      else:
         fractions_ext = "tiff"
         pattern = "*Fractions.tiff"
      ```        
8. The code assumes that you have a pixel size table named pixelsizes.txt located at the location of the base root. There is an example file provided here that you can modify. You can omit the beam size column for non-Tundra microscopes.

# Running the Script
  ```bash
   cd /path/to/screening_vis
   conda activate screening_vis
   python app.py
   ```
Copy and paste the http://XXX.X.X.X:XXXX into the web browser on the support PC and it should allow you to interactively browse screening and data collection sessions
