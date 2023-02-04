

class Guitar:
    def __init__(self):
        self.n_strings=6 # attribute
        self.Play()      # faccio partire in automatico il metodo Play()
    def Play(self): # Se non aggiungo self non ho un metodo ma una funzione
                    # Possiamo chiamare un metodo dopo aver creato un oggetto
                    # MA
                    # possiamo chiamare un metodo all'interno di una Classe inserendolo in __init__
                    # Sarà automaticamente eseguito ogni volta che creo un oggetto
            print('do re mi fa sol')

# class Electric_Guitar:
#     def __init__(self):
#         self.n_strings=6 # attribute
#         self.is_distrorsion = 'True' #solo per questa riga non devo ripetere tutto
#         self.Play() # faccio partire in automatico il metodo Play()
#
#     def Play(self): # Se non aggiungo self non ho un metodo ma una funzione
#                     # Possiamo chiamare un metodo dopo aver creato un oggetto
#                     # MA
#                     # possiamo chiamare un metodo all'interno di una Classe inserendolo in __init__
#                     # Sarà automaticamente eseguito ogni volta che creo un oggetto
#             print('do re mi fa sol')

# OPPURE
class Electric_Guitar(Guitar):
    pass # Non vuoi fare nessun cambiamento nella Classe
         # Corrispondono al 100%





# Per accedere a n_strings
# Prima devi creare a Guitar object

my_guitar          = Guitar()
my_electric_guitar = Electric_Guitar()

# Chiamiamo l'attributo o il metodo su un oggetto
# print(my_guitar.n_strings) #oggetto.variabile NO()
# print(my_guitar.Play())    #oggetto.metodo()

print(my_electric_guitar.n_strings) #Eredito tutto