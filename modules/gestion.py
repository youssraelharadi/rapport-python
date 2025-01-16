class GestionTaches:
    def __init__(self):
        self.taches = []

    def ajouter_tache(self, tache):
        self.taches.append(tache)

    def afficher_taches(self):
        for idx, tache in enumerate(self.taches, 1):
            print(f"{idx}. {tache}")