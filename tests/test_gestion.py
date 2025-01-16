import unittest
from modules.gestion import GestionTaches

class TestGestionTaches(unittest.TestCase):
    def test_ajouter_tache(self):
        gestion = GestionTaches()
        gestion.ajouter_tache("Apprendre Python")
        self.assertEqual(len(gestion.taches), 1)
        self.assertEqual(gestion.taches[0], "Apprendre Python")

if _name_ == "_main_":
    unittest.main()