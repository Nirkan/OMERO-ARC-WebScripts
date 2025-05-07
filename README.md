Note : These are some custom scripts used in the [presentation](https://zenodo.org/records/15308773) (version1) . In most cases offical (import/export) scripts in OMERO are recommended, however conversion of .xlsx to .csv file format might be needed for OMERO and ARC integration.

### About
Scripts to transfer metadata between OMERO and ARC via OMERO.web scripts. Import ISA metadata from ARC to OMERO and key-value pairs from OMERO to ARC as excel. The Regions of Interest to and from OMERO as json file. This metadata can be used to populate the ARC or populate OMERO with metadata relate to the investigation, study and assay as well as the Regions of Interst(ROIs) metadata. Metadata regarding each individual image in the dataset can exported from OMERO into excel rows and from excel rows to OMERO.

## Installation

The the following packages need to be installed by the system administrator.
pandas
openpyxl
