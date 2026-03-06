# EPU_Screening_Visualization
A web app using Flask that can run on a support PC and allow users to visualize screening and collection sessions running in a single particle session on EPU. Designed for use with a Tundra / Ceta-F setup, but should be compatible with other microscope/camera systems running EPU.

# Prerequisites

1. This app requires that Windows Subsystem for Linux (WSL) and miniconda are downloaded on the support PC. All commands should be run in the WSL terminal. You should have the output files mounted on the a drive accessible via /mnt/z (or your letter of choice). If your drive isn't mounted, you can run the following to mount it to z:
```bash
sudo mount -t drvfs Z: /mnt/z
```

2. The app assumes the following directory structure (standard for EPU) aside from the mount location. If you change this structure, you may also have to modify the script accordingly.
   
<img width="248" height="260" alt="image" src="https://github.com/user-attachments/assets/4b8f00cf-b00f-47f6-bcf1-90aedb0912fb" />

# Getting Started

3. Generate a conda environment appropriate for the screening visualization
   ```bash
   conda create --name screening_vis python=3.9 reportlab Flask pandas numpy Pillow
   ```
4. Activate the conda environment
   ```bash
   conda activate screening_vis
   ```
5. Modify this section of app.py if you have the drive mounted in a different location or if you have pixel sizes located somewhere other than in the base directory (see step 7 for more on this)
   
   ```python
   BASE_ROOT = "/mnt/z"
   PIXEL_TABLE_PATH = os.path.join(BASE_ROOT, "pixelsizes.txt")
   ```
   
6. Modify this section of epu/epustats.py with your microscope information

   ```python
   MICROSCOPE_INFO = {
       "TUNDRA-XXX": ("DFCI Tundra", 1.6),
       "TITANXXX": ("HMS Krios2", 2.7),
       "TITANXXX": ("HMS Krios1", 2.7),
   }

   windows_root = "Z:\\"
   ```

   MICROSCOPE_INFO should contain your InstrumentModel (replace XXX with the serial number) followed by the
   spherical aberration in mm. If you do not know what to use for InstrumentModel, you can run 
   ```bash
   grep InstrumentModel FoilHole*.xml
   ```
   in a bash terminal within Images-Disc1/GridSquare*/Data/ (repalace the * after FoilHole with a single file
   name)
   
   The windows_root is the root of where the atlas is stored. This is most likely the same drive as your
   base_root above, but as it appears in the actual .xml files (usually you can just change the
   letter in the format above -- there should be one extra backslash as shown above). If you are not sure what
   your root is, you can run
   ```bash
   grep -oP 'AtlasId .{0,50}' EpuSession.dm
   ```
   in the terminal while standing in the base directory of an imaging session. This is only used to "clean up"
   the atlas path displayed in the summary table.
   
7. The code assumes that you have a pixel size table named pixelsizes.txt located at the location of the base root (where your EPU sessions are written, NOT where the script is located). There is an example file provided here that you can modify. You can omit the beam size column for 3-condenser systems.

8. If your microscope writes out .tiff files, or if you have a Ceta-F that writes out .mrc files, you do
   not need to change anything. Otherwise, you will have to modify this segment of epu/epustats.py to include
   your file extension(s). 
   ```python
    if cam_name == "Ceta-F":
        fractions_ext = "mrc"
        pattern = "*Fractions.mrc"
    else:
        fractions_ext = "tiff"
        pattern = "*Fractions.tiff"
   ```

# Running the Script
  ```bash
   cd /path/to/screening_vis
   conda activate screening_vis
   python app.py
   ```
Copy and paste the http://XXX.X.X.X:XXXX into the web browser on the support PC and it should allow you to interactively browse screening and data collection sessions

_The folders queried must be the full EPU-generated folders containing all metadata_

Note 1: You could definitely combine integrate epustats.py into generate_report.py, it is kept separate only because of our transfer workflow

Note 2: generate_report.py should be able to automatically find the atlas images if the directory is located within the screening/collection directory or the directory containing it. However you can provide the path to the directory containing the atlas images if it is located elsewhere. If there are no atlas images, the report will still generate but will skip showing atlas images. This version of the script does not find atlas images for non-Tundra data sets but could be easily modified to do so. 

_These scripts were generated with the assistance of GPT4DFCI, a private, HIPAA-secure endpoint to GPT-4o provided by DFCI_

# Example Screenshots
<img width="607" height="342" alt="screenshot1" src="https://github.com/user-attachments/assets/55dc7a57-a4cc-49b3-b28e-df1134f49688" />
<img width="935" height="914" alt="screenshot2" src="https://github.com/user-attachments/assets/77588eaa-66b7-47a9-9660-0ff1365b3ebd" />
<img width="894" height="982" alt="screenshot3" src="https://github.com/user-attachments/assets/0215886f-1f5c-444a-b086-80c85c81ea68" />
<img width="839" height="528" alt="screenshot4" src="https://github.com/user-attachments/assets/94e268b3-62f6-4325-9477-ee4f70851015" />
<img width="1457" height="713" alt="screenshot5" src="https://github.com/user-attachments/assets/23dacb19-880d-48a7-8bdd-9e09681a9a5d" />
