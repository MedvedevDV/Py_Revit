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
        
        
mark = Selections.pick_element_by_class(IndependentTag)
if mark:
    el = doc.GetElement(mark.TaggedElementId.HostElementId).Category
    middle = mark.LeaderElbow
    head = mark.TagHeadPosition
    point = mark.LeaderEnd


    taggets = Selections.pick_elements_by_class(IndependentTag)
    with Transaction(doc, "Выровнять марки") as t:
        t.Start()
        for tag in taggets:
            tag.TagHeadPosition = head
            tag.LeaderElbow = middle
        t.Commit()
