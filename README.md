# DePipe_Demo

This project includes parts of modules used in the complete automated deep excavation modeling system

The modules offered here are:

**1. Data reader**\
    Transform the design data (excel) to parameters for modeling in advance.

**2. Manhole and pipe modeling**\
    Use the coodinates to create different types of manholes based on data. Similarly, continuous coordinates forms the pipe models. 
 
**3. Section creating**\
    The sections of pipes have lots of different parameters like diameters, thickness and size(3x4,4x4...).
    This module use a parametric section model to create the needed ones based on the data of projects. Then, it will be swept in modeling module.
  
**4. Upload to BIM360**\
    Create a web browser in the Winform application of project with CefSharp for logging in Forge.
    Send the request to get the token to log in Forge, then choose the model to be saved in the cloud storage which is connected with BIM360.

### Demo
https://youtu.be/TAzdiduCY_I

### Image

![image](https://github.com/Vincentaee/DePipe_Demo/blob/master/section2pipe.jpg)
section to pipe

![image](https://github.com/Vincentaee/DePipe_Demo/blob/master/forgeUpload.jpg)
upload to forge and BIM360
