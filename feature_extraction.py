""" Loads a STEP file and identify geometrical nature of each face
(cylindrical face, planar etc.)
See github issue https://github.com/tpaviot/pythonocc-core/issues/470

Two options in this example:

1. Click any planar or cylindrical face from the 3d window. They will
be identified as known surfaces, their properties displayed in the
console

2. A batch mode : click the menu button. All the faces will be traversed
and analyzed
"""

from __future__ import print_function

import os
import os.path
import sys
import json
import logging
from GraphManage import searchProcess

from OCC.Core.STEPControl import STEPControl_Reader
from OCC.Core.IFSelect import IFSelect_RetDone, IFSelect_ItemsByEntity
from OCC.Core.GeomAbs import GeomAbs_Plane, GeomAbs_Cylinder
from OCC.Core.TopoDS import topods_Face
from OCC.Core.BRepAdaptor import BRepAdaptor_Surface
from OCC.Display.SimpleGui import init_display
'''
from occ.Core.STEPControl import STEPControl_Reader
from occ.Core.IFSelect import IFSelect_RetDone, IFSelect_ItemsByEntity
from occ.Core.GeomAbs import GeomAbs_Plane, GeomAbs_Cylinder
from occ.Core.TopoDS import topods_Face
from occ.Core.BRepAdaptor import BRepAdaptor_Surface
from occ.Display.SimpleGui import init_display
from occ.Extend.TopologyUtils import TopologyExplorer
from pyautogui import hotkey
'''

def read_step_file(filename):
    """read the STEP file and returns a compound"""
    step_reader = STEPControl_Reader()
    status = step_reader.ReadFile(filename)

    if status == IFSelect_RetDone:  # check status
        failsonly = False
        step_reader.PrintCheckLoad(failsonly, IFSelect_ItemsByEntity)
        step_reader.PrintCheckTransfer(failsonly, IFSelect_ItemsByEntity)
        step_reader.TransferRoot(1)
        a_shape = step_reader.Shape(1)
    else:
        print("Error: can't read file.")
        sys.exit(0)
    return a_shape


index_cy = 0
index_pl = 0
features = {}
features["planes"] = {}
features["cylinders"] = {}
planes = []
cylinders = []


def recognize_face(a_face):
    """Takes a TopoDS shape and tries to identify its nature
    whether it is a plane a cylinder a torus etc.
    if a plane, returns the normal
    if a cylinder, returns the radius
    """
    global index_pl, index_cy
    surf = BRepAdaptor_Surface(a_face, True)
    surf_type = surf.GetType()
    if surf_type == GeomAbs_Plane:
        # print("--> plane " + str(index_pl))
        # look for the properties of the plane
        # first get the related gp_Pln
        gp_pln = surf.Plane()
        location = gp_pln.Location()  # a point of the plane
        # gp_pln.
        normal = gp_pln.Axis().Direction()  # the plane normal
        # then export location and normal to the console output
        # print(
        #     "--> Location (global coordinates)",
        #     round(location.X(), 2),
        #     round(location.Y(), 2),
        #     round(location.Z(), 2),
        # )
        # print("--> Normal (global coordinates)", round(normal.X(), 2), round(normal.Y(), 2), round(normal.Z(), 2))
        test_plane = (location.X(), location.Y(), location.Z(), normal.X(), normal.Y(), normal.Z())
        new_plane = {"location": (round(location.X(), 2), round(location.Y(), 2), round(location.Z(), 2)),
                     "normal": (round(normal.X(), 2), round(normal.Y(), 2), round(normal.Z(), 2))}
        if test_plane not in planes:
            index_pl += 1
            planes.append(test_plane)
            features["planes"][index_pl] = new_plane
    elif surf_type == GeomAbs_Cylinder:
        # print("--> cylinder " + str(index_cy))
        # look for the properties of the cylinder
        # first get the related gp_Cyl
        gp_cyl = surf.Cylinder()
        location = gp_cyl.Location()  # a point of the axis
        axis = gp_cyl.Axis().Direction()  # the cylinder axis
        # then export location and normal to the console output
        # print(
        #     "--> Location (global coordinates)",
        #     round(location.X(), 2),
        #     round(location.Y(), 2),
        #     round(location.Z(), 2),
        # )
        # print("--> Axis (global coordinates)", axis.X(), axis.Y(), axis.Z())
        test_cylinder = (location.X(), location.Y(), location.Z(), axis.X(), axis.Y(), axis.Z())
        new_cylinder = {"location": (round(location.X(), 2), round(location.Y(), 2), round(location.Z(), 2)),
                        "axis": (round(axis.X(), 2), round(axis.Y(), 2), round(axis.Z(), 2))}
        if test_cylinder not in cylinders:
            index_cy += 1
            cylinders.append(test_cylinder)
            features["cylinders"][index_cy] = new_cylinder
    else:
        # TODO there are plenty other type that can be checked
        # print(surf_type)
        # see documentation for the BRepAdaptor class
        # https://www.opencascade.com/doc/occt-6.9.1/refman/html/class_b_rep_adaptor___surface.html
        print("not implemented")
    return surf_type


def recognize_clicked(shp, *kwargs):
    """This is the function called every time
    a face is clicked in the 3d view
    """
    for shape in shp:  # this should be a TopoDS_Face TODO check it is
        # print("Face selected: ", shape)
        print('\n' * 20)
        res = recognize_face(topods_Face(shape))
        if res == GeomAbs_Plane:
            print("该特征为平面，加工流程如下：")
            searchProcess('平面', 6, 0.08)
        elif res == GeomAbs_Cylinder:
            print("该特征为孔，加工流程如下：")
            searchProcess('孔', 6, 0.08)


def recognize_batch(event=None):
    """Menu item : process all the faces of a single shape"""
    # then traverse the topology using the Topo class
    t = TopologyExplorer(shp)
    # t.faces_from_edge()cam_parameters.json
    # loop over faces only
    print("hi")
    print(f"number of vertices: {t.number_of_vertices()}")
    print(f"number of faces: {t.number_of_faces()}")
    print(f"number of wires: {t.number_of_wires()}")
    print("hello")
    for f in t.faces():
        # call the recognition function
        recognize_face(f)
        print("\n")
    # print(len(planes))
    # print(len(cylinders))
    with open("./features.json", "w") as f:
        json.dump(features, f, indent=4, ensure_ascii=False)
        # json.dump(cylinders, f, indent=4, ensure_ascii=False)
    print("File written")


def exit(event=None):
    sys.exit()


if __name__ == "__main__":
    logging.basicConfig(filename='logger.log', level=logging.INFO)
    display, start_display, add_menu, add_function_to_menu = init_display()
    display.SetSelectionModeFace()  # switch to Face selection mode
    display.register_select_callback(recognize_clicked)
    # first loads the STEP file and display
    shp = read_step_file(".\object2.STEP")
    display.DisplayShape(shp, update=True)
    add_menu("recognition")
    add_function_to_menu("recognition", recognize_batch)
    start_display()
