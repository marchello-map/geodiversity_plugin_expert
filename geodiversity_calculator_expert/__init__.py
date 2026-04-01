# -*- coding: utf-8 -*-
"""GeodiversityCalculator Expert v0.99"""

def classFactory(iface):
    from .geodiversity_calculator import GeodiversityCalculator
    return GeodiversityCalculator(iface)
