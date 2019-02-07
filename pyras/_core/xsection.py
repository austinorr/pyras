
from . import Node

class CrossSection(Node):
    """
    """

    def __init__(self):
        """

        """
        self.banks

    # Alternate constructors
    @staticmethod
    def from_coord(x, y):
        """

        """    
    
    def reverse(self):
        """
        """
    
    # Hydraulic properties
    def area(self, z=None, y=None):
        """

        """
    
    def perimeter(self, z=None):
        """

        """
        
    def radius(self, z=None):
        """

        """
    
    def conveyance(self, z=None):
        """

        """

    @property
    def banks(self):
        """
        """

    @banks.setter
    def banks(self, left, right):
        """
        """

    @property
    def angle(self):
        """
        """

    @banks.setter
    def angle(self, value):
        """
        """

    @property
    def offset(self):
        """
        """

    @banks.setter
    def offset(self, y):
        """
        """
        
        
    def conveyance(self, flow, z=None, y=None):
        """
        """
        
    def plot(self):
        """
        """
