#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Export image metadata and key-value annotations from a Dataset or Image list to an Excel file.
The result is attached as a downloadable OMERO file annotation.
"""

import omero
from omero.gateway import BlitzGateway, MapAnnotationWrapper
from omero.rtypes import rstring, robject
from omero import scripts
import pandas as pd
import tempfile
import os


def extract_image_metadata(images):
    data = []
    for image in images:
        image_name = image.getName()
        X = image.getSizeX()
        Y = image.getSizeY()
        Z = image.getSizeZ()
        C = image.getSizeC()
        T = image.getSizeT()

        kv_pairs = {}
        for ann in image.listAnnotations():
            if isinstance(ann, MapAnnotationWrapper):
                for key, value in ann.getValue():
                    kv_pairs[key] = value

        row = {
            "ImageName": image_name,
            "PixelSizeX": X,
            "PixelSizeY": Y,
            "PixelSizeZ": Z,
            "Channels": C,
            "TimeAxis": T,
        }
        row.update(kv_pairs)
        data.append(row)
    return data


def run_script():
    client = scripts.client(
        "Export_Image_Metadata_Excel.py",
        "Export image metadata and key-value pairs to Excel, with download link.",
        scripts.String("Data_Type", optional=False, values=["Dataset", "Image"],
                       description="Select type of input: Dataset or Image"),
        scripts.List("IDs", optional=False, description="List of Dataset or Image IDs").ofType(scripts.rlong(0))
    )

    try:
        params = client.getInputs(unwrap=True)
        data_type = params["Data_Type"]
        ids = params["IDs"]

        conn = BlitzGateway(client_obj=client)
        
        images = []
        if data_type == "Dataset":
            for ds_id in ids:
                dataset = conn.getObject("Dataset", ds_id)
                if dataset:
                    images.extend(list(dataset.listChildren()))
        else:
            for img_id in ids:
                image = conn.getObject("Image", img_id)
                if image:
                    images.append(image)

        if not images:
            client.setOutput("Message", rstring("No images found."))
            return

        data = extract_image_metadata(images)
        df = pd.DataFrame(data)

        # Save Excel file to a temporary location
        temp_dir = tempfile.mkdtemp()
        filename = "ImageToFile.xlsx"
        filepath = os.path.join(temp_dir, filename)
        df.to_excel(filepath, index=False)

        # Upload the file as a downloadable OMERO file annotation
        file_ann = conn.createFileAnnfromLocalFile(
            filepath,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ns="omero.script_results",
            desc="Exported image metadata as Excel"
        )

        client.setOutput("File_Annotation", robject(file_ann._obj))
        client.setOutput("Message", rstring("Click the file above to download."))

    finally:
        client.closeSession()


if __name__ == "__main__":
    run_script()

