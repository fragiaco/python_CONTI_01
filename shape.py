

class Rectangle:
    def __init__(self, height, width):
        self.height = height
        self.width = width

    def area_rettangolo(self):
        return self.height * self.width


# class Square:
#     def __init__(self, height):
#         self.height = height
#
#     def area(self):
#         return self.height * self.height

class Square(Rectangle):
    def __init__(self, height):          #fornisco un lato
        super().__init__(height, height)      #fornisco a height = l e width = l

    def perimetro_quadrato(self):
        return self.height * 4

# r1 = Rectangle(2, 5)
# print(r1.area())
#
# Sq = Square(2)
# print(Sq.area())


class Cube(Square):
    def surface_area(self):
        face_area = self.area_rettangolo()
        return face_area * 6

    def volume(self):
        face_area = super().area_rettangolo()
        return face_area * self.height

C = Cube(3)
print(C.surface_area())
print(C.volume())