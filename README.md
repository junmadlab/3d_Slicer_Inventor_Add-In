# 3d Slicer: An Inventor Add-In Using Visual Basic

## Information

This Add-In, which is developed on Inventor API using Visual Basic, is an automated 3D printing slicer targeted for Stereolithography (SLA) processes. Compared to the existing SLA slicers, rather than (1) generating only the 2D image for every slice, this slicer can also help (2) obtain the 3D solid of each slice or (3) the currently cured part stacked by all previous layers. The cross-section 2D image of each slice is saved with a ".STL" extension (easy to be converted into .PNG if needed), and the 3D solids are stored as STEP CAD files with a ".stp" extension, which can be conveniently used for further FE simulations or other operations.

This Add-In is developed to create the database for [the paper](https://link.springer.com/article/10.1007/s00170-019-03363-4) "*Deep Learning-Based Stress Prediction for Bottom-Up SLA 3D Printing Process*. The International Journal of Advanced Manufacturing Technology 102, no. 5-8 (2019): 2555-2569." For the specific application, please refer to the paper.

## Usage

### Scripts

* Slicing for cross-section 2D images: "crossSection_v1.bas" and "crossSection_v0.bas".

  * "crossSection_v1.bas": The final-version script for generating cross-section 2D images.
  * "crossSection_v0.bas": This script will first move the part so that its centroid of the bottom is coincident with the origin, and then generate cross-section 2D images.
  
* Slicing for 3D solids (currently cured part stacked by all previous layers by default): "slicing_v1.bas" and "slicing_v0.bas".
  
  * "slicing_v1.bas": The final-version script for generating 3D solids of currently cured part stacked by all previous layers. In order to obtain 3D solids of each slice, the user can easily modify the script to realize that.
  * "slicing_v0.bas": This script will first move the part so that its centroid of the bottom is coincident with the origin, and then generate 3D solids.
  
### Run Scripts

1. Save the parts (.stp) for slicing in a directory (e.g., "C:\PATH\TO\PARTS\").
2. Open Inventor (The Add-In is developed using Inventor 2018. Later versions will work fine).
3. Go to **Tools** and click "VBA Editor" (A new window named "Microsoft Visual Basic for Applications" will appear).
4. Go to **File -> Import File** and load the VBA script (e.g., "crossSection_v1.bas").
5. A "Module #" will be loaded and the VBA code will be shown in the window.
6. Change Line 11 (strDir = "C:\PATH\TO\PARTS\") to the directory you saved the parts (.stp) for slicing. Change Line 86 (n = 10) to the number of layers you want to slice. Change Line 328 (Const STLFilePath As String = "C:\PATH\TO\STL\IMAGES") to the directory you want to save the sliced cross-section 2D images (.stl). Similar steps for the VBA script "slicing_v1.bas".
7. Click **Run Macro** to run the script and all the parts saved in the directory will be automatically sliced. The cross-section 2D images and generated 3D solids will also be automatically saved in the specified directories.

## Cite

Please cite [our paper](https://link.springer.com/article/10.1007/s00170-019-03363-4) (and the respective papers of the methods used) if you use this code in your own work:
```
@article{khadilkar2019deep,
  title={Deep learning--based stress prediction for bottom-up SLA 3D printing process},
  author={Khadilkar, Aditya and Wang, Jun and Rai, Rahul},
  journal={The International Journal of Advanced Manufacturing Technology},
  volume={102},
  number={5-8},
  pages={2555--2569},
  year={2019},
  publisher={Springer}
}
```
