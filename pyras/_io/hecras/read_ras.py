"""
"""


class SimpleAttribute:
    """
    """    
    def __init__(self, name, options=['SI Units', 'English Units']):
        pass


class NamedAttribute:
    """
    """    
    def __init__(self, name, type_, value=None, separator='=',
                 max_length=None):
        pass


class TagAttribute:
    """
    """
    def __init__(self, name, start_tag, end_tag, type_, value=None,
                 max_length=None):
        pass



def _generic_reader():
    """ """


def read_project(filename):
    """
Proj Title=new_project
Default Exp/Contr=0.3,0.1
SI Units
Y Axis Title=Elevation
X Axis Title(PF)=Main Channel Distance
X Axis Title(XS)=Station
BEGIN DESCRIPTION:
Example text
END DESCRIPTION:
DSS Start Date=
DSS Start Time=
DSS End Date=
DSS End Time=
DSS Export Filename=
DSS Export Rating Curves= 0 
DSS Export Rating Curve Sorted= 0 
DSS Export Volume Flow Curves= 0 
DXF Filename=
DXF OffsetX= 0 
DXF OffsetY= 0 
DXF ScaleX= 1 
DXF ScaleY= 10 
GIS Export Profiles= 0 
     
    """
    sep ='='
    
    tags = {
        'description': ['BEGIN DESCRIPTION:', 'END DESCRIPTION:']
        }

    fixed = {
        'units': ['SI Units' 'English Units']
        }

    keys = {
        'Proj Title': '',
        'Default Exp/Contr': '=0.3,0.1',
        'Current Plan': '=p03',
        'Geom File': '=g01',
        'Flow File': '=f01',
        'Plan File': '=p01',
        'Y Axis Title=Elevation': '',
        'X Axis Title(PF)': '=Main Channel Distance',
        'X Axis Title(XS)': '=Station',
        'DSS Start Date': '=',
        'DSS Start Time': '=',
        'DSS End Date': '=',
        'DSS End Time': '=',
        'DSS Export Filename': '=',
        'DSS Export Rating Curves': '= 0',
        'DSS Export Rating Curve Sorted': '= 0',
        'DSS Export Volume Flow Curves': '= 0',
        'DXF Filename': '=',
        'DXF OffsetX': '= 0',
        'DXF OffsetY': '= 0',
        'DXF ScaleX': '= 1',
        'DXF ScaleY': '= 10',
        'GIS Export Profiles': '= 0'
        }




def read_geometry(filename):
    """ """


def read_plan(filename):
    """ """


def read_boundary(filename):
    """ """


def test_project():
    """ """
    'D:\Users\penac1\Dropbox (Personal)\it\repos\git\pyras\temp_examples\Steady Examples'


if __name__ == '__main__':
    test_project()
