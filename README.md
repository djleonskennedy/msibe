msibe
=====

allows to extrat binary data to file

MSI Binary Extractor (Transform.mst)

This code should be placed into MSI database Custom Action (immediate mode)

msi custom action:

Custom Action Table: ExtractBinary	38  Code From "Extract.vbs"

InstallExecuteTable: ExtractBinary  NOT REMOVE  1800
