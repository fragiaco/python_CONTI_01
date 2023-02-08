

from shape import Square

# Accedere dalla Child class (Square()) dalla Parent class (modulo area())

Sq = Square(2)
print(f"L'area del quadrato è = {Sq.area_rettangolo()}")
print(f"Il perimetro del quadrato è = {Sq.perimetro_quadrato()}")

# Uso super() per accedere dalla Parent class (Rectengle()) alla child class (modulo perimeter())
from shape import Cube
C = Cube(2)
print(f"L'area della superficie di un lato del cubo è = {super(Cube, C).area_rettangolo()}")
#non surface_area perchè guarda nelle Classi da Cubo escluso in su
print(f"Il perimetro del rettangolo è = {super(Cube,C).perimetro_quadrato()}")


