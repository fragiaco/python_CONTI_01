

class Component:

# composite class constructor
	def __init__(self):
		print('Component class object created...')

	# composite class instance method
	def m1(self):
		print('Component class m1() method executed...')


class Composite:

	# composite class constructor
	def __init__(self):

		# creating object of component class
		self.obj1 = Component()

		print('Composite class object also created...')

	# composite class instance method
	def m2(self):

		print('Composite class m2() method executed...')

		# calling m1() method of component class
		self.obj1.m1()


# creating object of composite class
obj2 = Composite()

# calling m2() method of composite class
obj2.m2()



class Giacomo_Component:
	def __init__(self):
		print('Component:')

	def mi_chiamo_giacomo(self):
		print('Mi chiamo giacomo')

class Principale:
	def __init__(self):
		self.giacomo = Giacomo_Component()

#	def print_name(self):
		self.giacomo.mi_chiamo_giacomo()


obj = Principale()

