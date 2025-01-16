from modules.gestion import GestionTaches

if __name__== "__main__":
    gestion = GestionTaches()
    gestion.ajouter_tache("Apprendre Python")
    gestion.afficher_taches()