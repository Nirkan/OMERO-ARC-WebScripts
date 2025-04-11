# coding=utf-8
'''
OMERO Web Script: Upload metadata from an ISA Excel file as key-value pairs.

This script:
- Allows users to specify a File Annotation ID for an Excel file.
- Extracts metadata and stores it as OMERO key-value pairs.
- Attaches metadata to a selected Project, Dataset, or Screen.
'''

import omero
import omero.scripts as scripts
import omero.gateway
from omero.gateway import BlitzGateway, MapAnnotationWrapper
from omero.rtypes import rstring, rlong, robject

import pandas as pd
import io

def extract_metadata_from_xlsx(file_obj):
    """Extract metadata from an Excel file-like object using pandas."""
    metadata = {}
    try:
        xls = pd.ExcelFile(file_obj)
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name, dtype=str, header=None).fillna("")
            file_type = sheet_name
            namespace = None

            for _, row in df.iterrows():
                if row.iloc[0].strip().isupper():  # uppercase headers as namespace
                    namespace = f"ARC:ISA:{file_type.upper()}:{row.iloc[0].strip()}"
                    metadata[namespace] = {}
                elif namespace:
                    key = row.iloc[0].strip()
                    values = [str(v).strip() if str(v).strip() else '' for v in row[1:]]
                    if key:
                        metadata[namespace][key] = ", ".join(values)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
    return metadata

def apply_metadata(obj, metadata, conn):
    """Apply metadata as key-value pairs to an OMERO object."""
    for namespace, values in metadata.items():
        map_ann = MapAnnotationWrapper(conn)
        map_ann.setNs(namespace)
        map_ann.setValue(list(values.items()))
        map_ann.save()
        obj.linkAnnotation(map_ann)

def run_script():
    """Main entry point for the OMERO script."""
    client = scripts.client(
        "ISA_Metadata_Import",
        "Import metadata from an ISA Excel file attached as a FileAnnotation.",
        scripts.String("Object_Type", grouping="1", optional=False, values=["Project", "Dataset", "Screen"],
                      description="The type of object to annotate (Project, Dataset, or Screen)."),
        scripts.List("IDs", grouping="2", optional=False, description="Target object ID(s)").ofType(rlong(0)),
        scripts.String("File_Annotation", grouping="3", optional=False,
                       description="File Annotation ID of Excel file")
    )

    try:
        # Safer input handling and connection
        params = client.getInputs(unwrap=True)
        obj_type = params["Object_Type"]
        ids = params["IDs"]
        ann_id = int(params["File_Annotation"])

        if not ids:
            client.setOutput("Error", rstring("No IDs provided."))
            return

        obj_id = ids[0]
        conn = BlitzGateway(client_obj=client)
        obj = conn.getObject(obj_type, obj_id)

        if not obj:
            client.setOutput("Error", rstring(f"Invalid {obj_type} ID: {obj_id}"))
            return

        file_ann = conn.getObject("FileAnnotation", ann_id)
        if not file_ann:
            client.setOutput("Error", rstring(f"No FileAnnotation found with ID {ann_id}"))
            return

        file_stream = io.BytesIO()
        for chunk in file_ann.getFileInChunks():
            file_stream.write(chunk)

        file_stream.seek(0)
        metadata = extract_metadata_from_xlsx(file_stream)

        if not metadata:
            client.setOutput("Message", rstring("No metadata extracted from Excel file."))
            return

        apply_metadata(obj, metadata, conn)

        client.setOutput("Message", rstring(f"Metadata imported for {obj_type} ID {obj_id}."))

    finally:
        client.closeSession()

if __name__ == "__main__":
    run_script()

