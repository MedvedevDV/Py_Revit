# -*- coding: utf-8 -*-
import clr
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Collections")

import System
clr.AddReference("RevitAPI")
import Autodesk

clr.AddReference("RevitAPIUI")

import re
import sys
import System.Windows.Forms
from clr import StrongBox
from System.Collections.Generic import IList
from math import fabs

from Autodesk.Revit.DB import Dimension, Transaction, XYZ, BuiltInCategory
from Autodesk.Revit.DB import Reference, IndependentTag, Transform, Line
from Autodesk.Revit.DB import IntersectionResultArray, Plane
from Autodesk.Revit.DB import ClosestPointsPairBetweenTwoCurves

from Autodesk.Revit.DB import Category, BuiltInCategory
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
active_view = doc.ActiveView

#--------------Класс програмного выбора обьектов------------------------
class Get_revit_elements:
    """Класс для поиска элементов в Revit."""

    @classmethod
    def get_elems_by_category(cls, category_class, active_view=None, name=None):
        """Получение элемента по классу категории."""
        if not active_view:
            els = FilteredElementCollector(doc).OfClass(category_class).\
                  ToElements()
        else:
            els = FilteredElementCollector(doc, active_view).\
                  OfClass(category_class).ToElements()
        if name:
            els = [i for i in els if name in i.Name] 
        return els

    @classmethod
    def get_elems_by_builtinCategory(cls, built_in_cat=None, include=[],
                                     active_view=None):
        """Получение элемента по встроенному классу."""
        if not include:
            if not active_view:
                els = FilteredElementCollector(doc).OfCategory(built_in_cat)
            else:
                els = FilteredElementCollector(doc, active_view).\
                      OfCategory(built_in_cat)
            return els.ToElements()
        if include:
            new_list = []
            for i in include:
                if not active_view:
                    els = FilteredElementCollector(doc).OfCategory(built_in_cat)
                else:
                    els = FilteredElementCollector(doc, active_view).\
                          OfCategory(built_in_cat)
                new_list += els.ToElements()
            return new_list
        
#--------------Класс получения категории элемента-----------------------
class Pick_by_category(ISelectionFilter):
    doc = __revit__.ActiveUIDocument.Document
    def __init__(self, built_in_category):
        if isinstance(built_in_category, Category):
            self.built_in_category = [built_in_category.Id]
        else:
            if isinstance(built_in_category, BuiltInCategory):
                built_in_category = [built_in_category]
            self.built_in_category = [Category.GetCategory(self.doc, i).Id for i in built_in_category]

    def AllowElement(self, el):
        if el.Category.Id in self.built_in_category:
            return True
        return False

    def AllowReference(self, refer, xyz):
        return False
    
#--------------Класс получения класса элемента-----------------------
class Pick_by_class(ISelectionFilter):
    doc = __revit__.ActiveUIDocument.Document
    def __init__(self, class_type):
        self.class_type = class_type

    def AllowElement(self, el):
        if isinstance(el, self.class_type):
            return True
        return False

    def AllowReference(self, refer, xyz):
        return False

#--------------Класс ручного выбора элементов-----------------------
class Selections:
    """Класс с реализацией различных методов выбора элементов."""

    selection = __revit__.ActiveUIDocument.Selection
    doc = __revit__.ActiveUIDocument.Document
    @classmethod
    def pick_element_by_category(cls, built_in_category):
        """Выбор одного элемента по BuiltInCategory."""
        try:
            return cls.doc.GetElement(cls.selection.PickObject(ObjectType.Element, Pick_by_category(built_in_category)))
        except OperationCanceledException:
            return

    @classmethod
    def pick_element_by_class(cls, class_type):
        """Выбор одного элемента по категории."""
        try:
            return cls.doc.GetElement(cls.selection.PickObject(ObjectType.Element, Pick_by_class(class_type)))
        except OperationCanceledException:
            return

    @classmethod
    def pick_elements_by_class(cls, class_type):
        """Выбор одного элемента по категории."""
        try:
            return [cls.doc.GetElement(i) for i in cls.selection.PickObjects(ObjectType.Element, Pick_by_class(class_type))]
        except OperationCanceledException:
            return
        

def vect_cur_view(view):
    v_right = active_view.RightDirection
    v_right = XYZ(fabs(v_right.X), fabs(v_right.Y), fabs(v_right.Z))
    v_up = active_view.UpDirection
    v_up = XYZ(fabs(v_up.X), fabs(v_up.Y), fabs(v_up.Z))
    return (v_right + v_up)

def get_active_ui_view(uidoc):
    doc = uidoc.Document
    view = doc.ActiveView
    uiviews = uidoc.GetOpenUIViews()
    uiview = None
    for uv in uiviews:
        if uv.ViewId.Equals(view.Id):
            uiview = uv
            break
    return uiview

def get_coordinate():
    uiview = get_active_ui_view(uidoc)
    rect = uiview.GetWindowRectangle()
    p = System.Windows.Forms.Cursor.Position
    dx = float(p.X - rect.Left) / float(rect.Right - rect.Left)
    dy = float(p.Y - rect.Bottom) / float(rect.Top - rect.Bottom)
    v_right = active_view.RightDirection
    v_right = XYZ(fabs(v_right.X), fabs(v_right.Y), fabs(v_right.Z))
    v_up = active_view.UpDirection
    v_up = XYZ(fabs(v_up.X), fabs(v_up.Y), fabs(v_up.Z))
    dxyz = dx * v_right + dy * v_up

    corners = uiview.GetZoomCorners()

    a = corners[0]
    b = corners[1]
    v = b - a

    q = a + dxyz.X * v.X * XYZ.BasisX + dxyz.Y * v.Y * XYZ.BasisY + dxyz.Z * XYZ.BasisZ * v.Z
    return q

def SignedDistanceTo(plane, p):
    v = p - plane.Origin
    return plane.Normal.DotProduct(v)

def project_onto(plane, p):
    d = SignedDistanceTo(plane, p)
    q = p - d * plane.Normal
    return q

#--------------Начало-----------------------
mark = Selections.pick_element_by_class(IndependentTag)
if mark:
    el = doc.GetElement(mark.TaggedElementId.HostElementId).Category
    middle = mark.LeaderElbow
    head = mark.TagHeadPosition
    point = mark.LeaderEnd

if mark:
    plane = Plane.CreateByNormalAndOrigin(active_view.ViewDirection, active_view.Origin)
    middle = project_onto(plane, middle)
    head = project_onto(plane, head)
    point = project_onto(plane, point)

    tagget = True
    while tagget:
        tagget = Selections.pick_element_by_category(el)
        if tagget:
            pos = get_coordinate()
            ref = Reference(tagget)
            with Transaction(doc, 'Добавить выноску') as t:
                t.Start()
                vec = (point - middle).Normalize()
                n_pos = project_onto(plane, pos)
                l1 = Line.CreateUnbound(n_pos, vec)
                l2 = Line.CreateUnbound(middle, (head - middle).Normalize())
                inter = StrongBox[IntersectionResultArray]()
                l1.Intersect(l2, inter)
                new_middle = list(inter.Value)[0].XYZPoint
                new_mark = IndependentTag.Create(doc, mark.GetTypeId(), doc.ActiveView.Id, ref, True, mark.TagOrientation, head)
                new_mark.TagHeadPosition  = head
                new_mark.LeaderEndCondition  = mark.LeaderEndCondition
                new_mark.LeaderElbow = new_middle
                # new_pos = pos - middle
                new_mark.LeaderEnd  = pos
                if new_mark.TagText != mark.TagText:
                    doc.Delete(new_mark.Id)
                t.Commit()
