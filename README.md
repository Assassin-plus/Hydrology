# An automated python script to deal with hydrological data
## Usage
* 创建一个总文件夹作为工作环境。包括:
    * 一个`data`文件夹，用于存放原始数据
        *`data`文件夹中有`Boundary`文件夹存储shp类子流域边界数据，`global_ndvi`文件夹存储每年的水文学数据，`sand`文件夹存储土壤岩性数据，`Source`文件夹存储ohio河的遥感图
    * 一个`output`文件夹，用于存放处理后的数据
    * 一个`done`文件夹，用于存放处理脚本
以下文件路径请根据实际情况修改python脚本中的路径
* 将原始数据放入`data`文件夹中，包括
    * 一个`Source`文件夹，用于存放原始数据，如"ohio_90.tif"
    * 一个`Boundary`文件夹，用于存放边界掩膜数据，如"\\Ohio\\ohio_mask.shp"
* 使用GIS自带的python解释器运行脚本，即可得到处理后的数据
```python
C:\python27\ArcGIS10.2\python.exe {Analyze/coverage/ndvi/sand}.py
```
## 说明
因为ArcGIS的python库只能在ArcGIS自带的python解释器中使用，所以需要使用ArcGIS自带的python解释器运行脚本。如果需要在其他python解释器中运行，需要import脚本中的`arcpy`库。

ArcGIS的arcpy库内存泄漏，运行过程中如果数据太多，可能导致处理数据冲突，分批将数据放入Boundary处理即可。