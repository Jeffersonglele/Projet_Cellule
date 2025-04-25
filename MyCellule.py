import ttkbootstrap as ttk
import tkinter as tk
from ttkbootstrap.constants import *
from tkinter import filedialog as fd
from tkinter import messagebox
import csv
import pandas as pd


class MyCellule: #la classe principale qui gère l'interface graphique et les fonctionnalités
    def __init__(self, root): #Fenêtre principale de l'application.
        self.root = root
        self.root.title('Notre CELLULE')
        self.root.geometry("800x600")
        self.root.resizable(True, True)

        # Application du thème ttkbootstrap
        self.style = ttk.Style(theme="flatly") 

        # Liste pour stocker les références des cellules
        self.cellules = [] #Liste bidimensionnelle(double valeur) contenant les références des cellules

        # Historique des actions
        self.history = [] #Liste pour stocker l'historique des modifications des cellules
        self.index_actuel = -1

        # Délimiteur par défaut
        self.delimiteur = ','

        # Création des différents frames
        self.create_interface_graphique()
        self.create_bande()
        self.create_context_menu()

        # Initialisation de l'historique avec l'état vide
        self.mise_a_jour()

    def create_interface_graphique(self):
        """Crée une grille de cellules (21 lignes x 6 colonnes) avec des widgets Entry.
Associe les touches directionnelles (<Up>, <Down>, <Left>, <Right>) pour naviguer entre les cellules.
Configure le redimensionnement dynamique des lignes et colonnes."""
        # Frame pour contenir la grille
        self.grille_frame = ttk.Frame(self.root)
        self.grille_frame.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")

        # Dimensions de la grille
        nb_lignes = 21  # Nombre de lignes
        nb_colonnes = 6  # Nombre de colonnes

        # Créer une grille de cellules
        for i in range(nb_lignes):
            ligne = []
            for j in range(nb_colonnes):
                cellule = ttk.Entry(self.grille_frame, width=15, justify="center", bootstyle="info")
                cellule.grid(row=i, column=j, padx=1, pady=1)

                # Association des touches directionnelles
                cellule.bind("<Up>", lambda event, x=i, y=j: self.deplacer_focus(x - 1, y)) #haut
                cellule.bind("<Down>", lambda event, x=i, y=j: self.deplacer_focus(x + 1, y)) #bas
                cellule.bind("<Left>", lambda event, x=i, y=j: self.deplacer_focus(x, y - 1)) #gauche
                cellule.bind("<Right>", lambda event, x=i, y=j: self.deplacer_focus(x, y + 1)) #droit

                ligne.append(cellule)
            self.cellules.append(ligne)

        # Configuration du redimensionnement dynamique
        for i in range(nb_lignes):
            self.grille_frame.rowconfigure(i, weight=1)
        for j in range(nb_colonnes):
            self.grille_frame.columnconfigure(j, weight=1)

    def create_context_menu(self):
        """Crée un menu contextuel (clic droit) avec les options :Copier,Coller,Effacer et
associe le menu contextuel à chaque cellule"""
        self.menu_contextuel = tk.Menu(self.root, tearoff=0)
        self.menu_contextuel.add_command(label="Copier", command=self.copier_cellule)
        self.menu_contextuel.add_command(label="Coller", command=self.coller_cellule)
        self.menu_contextuel.add_command(label="Effacer", command=self.effacer_cellule)

        # Relation clic droit/cellules
        for row in self.cellules:
            for cell in row:
                cell.bind("<Button-3>", self.afficher_menu_contextuel)

    def afficher_menu_contextuel(self, event):
        """Affiche le menu contextuel"""
        # Stocker la cellule cliquée pour copier/coller
        self.cellule_active = event.widget
        self.menu_contextuel.post(event.x_root, event.y_root)

    def copier_cellule(self):
        """Copie le contenu de la cellule sélectionnée"""
        if hasattr(self, 'cellule_active'):
            self.clipboard = self.cellule_active.get()

    def coller_cellule(self):
        """Colle le contenu dans la cellule sélectionnée"""
        if hasattr(self, 'cellule_active') and self.clipboard:
            self.cellule_active.delete(0, tk.END)
            self.cellule_active.insert(0, self.clipboard)
            self.mise_a_jour()

    def effacer_cellule(self):
        """Efface le contenu de la cellule sélectionnée"""
        if hasattr(self, 'cellule_active'):
            self.cellule_active.delete(0, tk.END)
            self.mise_a_jour()

    def reorganiser_colonnes(self, event):
        """Réorganise les colonnes dynamiquement via glisser-déposer"""
        pass  #pas d'idée (fonctionnalité pas comprise)
    

    def deplacer_focus(self, x, y):
        """Déplace le focus vers la cellule à la position (x, y)"""
        if 0 <= x < len(self.cellules) and  0 <= y < len(self.cellules[x]):
            self.cellules[x][y].focus_set()

    def navigateur(self):
        filetypes = (
            ('Fichiers CSV', '*.csv'),
            ('Tous les fichiers', '*.*')
        )

        # Demande à l'utilisateur de sélectionner un fichier
        filename = fd.askopenfilename(filetypes=filetypes)

        if filename:
            try:
                with open(filename, 'r', encoding='utf-8') as f:
                    first_line = f.readline().strip()
                    if not first_line:
                        print("Le fichier est complètement vide.")  # Debug
                        messagebox.showerror("Erreur", "Le fichier CSV est complètement vide.")
                        return
                    reader = csv.reader(f, delimiter=self.delimiteur)

                    # Effacer toutes les cellules
                    self.nettoyeur()

                    # Remplir les cellules avec les données du CSV
                    data = []
                    for i, row in enumerate(reader):
                        if i >= 21:  # Limite de 21 lignes
                            break
                        row_data = []
                        for j, value in enumerate(row):
                            if j >= 5:  # Limite de 5 colonnes
                                break
                            self.cellules[i][j].delete(0, tk.END)
                            self.cellules[i][j].insert(0, value)
                            row_data.append(value)
                        data.append(row_data)

                    # Enregistrer l'état après chargement du fichier
                    self.history.append(data)
                    self.index_actuel = len(self.history) - 1
                    self.update_buttons()

            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de lire le fichier:\n{str(e)}")

    def nettoyeur(self):
        """Vide toutes les cellules"""
        for row in self.cellules:
            for cell in row:
                cell.delete(0, tk.END)

    def fusionner_fichiers_csv(self):
        """Fusionne plusieurs fichiers CSV dans les cellules"""
        filetypes = (
            ('Fichiers CSV', '*.csv'),
            ('Tous les fichiers', '*.*')
        )

        # Demande à l'utilisateur de sélectionner plusieurs fichiers
        filenames = fd.askopenfilenames(filetypes=filetypes)

        if filenames:
            try:
                self.nettoyeur()
                row_index = 0

                for filename in filenames:
                    with open(filename, 'r', encoding='utf-8') as f:
                        reader = csv.reader(f, delimiter=self.delimiteur)
                        for row in reader:
                            if row_index >= len(self.cellules):  # Limite de lignes
                                break
                            for col_index, value in enumerate(row):
                                if col_index >= len(self.cellules[row_index]):  # Limite de colonnes
                                    break
                                self.cellules[row_index][col_index].delete(0, tk.END)
                                self.cellules[row_index][col_index].insert(0, value)
                            row_index += 1

                # Enregistrer l'état après la fusion
                self.mise_a_jour()

            except FileNotFoundError:
                print("Erreur : Fichier introuvable.")  # Debug
                messagebox.showerror("Erreur", "Un ou plusieurs fichiers sont introuvables.")
            except UnicodeDecodeError:
                print("Erreur : Problème d'encodage.")  # Debug
                messagebox.showerror("Erreur", "Un ou plusieurs fichiers contiennent des caractères non valides ou un encodage incorrect.")
            except Exception as e:
                print(f"Erreur inattendue : {e}")  # Debug
                messagebox.showerror("Erreur", f"Impossible de fusionner les fichiers :\n{str(e)}")

    def fusionner_lignes(self):
        """Fusionne les valeurs des cellules sélectionnées dans une ligne"""
        try:
            for row in self.cellules:
                ligne_complete = " ".join([cell.get() for cell in row if cell.get()])
                for cell in row:
                    cell.delete(0, tk.END)
                if ligne_complete:
                    row[0].insert(0, ligne_complete)

            # Enregistrer l'état après la fusion
            self.mise_a_jour()

        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de fusionner les lignes:\n{str(e)}")

    def changer_delimiteur(self, delimiteur):
        """Change le délimiteur et recharge les données"""
        self.delimiteur = delimiteur
        self.navigateur()

    def mise_a_jour(self, event=None):
        """Enregistre l'état actuel des cellules"""
        etat_actuel = [[cell.get() for cell in row] for row in self.cellules]

        # Si l'état actuel est différent du dernier état enregistré
        if not self.history or etat_actuel != self.history[self.index_actuel]:
            # Supprimer les états après l'index courant
            self.history = self.history[:self.index_actuel + 1]
            # Ajouter le nouvel état
            self.history.append(etat_actuel)
            self.index_actuel = len(self.history) - 1
            self.update_buttons()

    def précédent(self):
        """Annule la dernière action"""
        if self.index_actuel > 0:
            self.index_actuel -= 1
            self.restore_state()
            self.update_buttons()

    def suivant(self):
        """Rétablit l'action annulée"""
        if self.index_actuel < len(self.history) - 1:
            self.index_actuel += 1
            self.restore_state()
            self.update_buttons()

    def restore_state(self):
        """Restaurer l'état depuis l'historique"""
        etat = self.history[self.index_actuel]
        
        for i, row in enumerate(etat):
            for j, value in enumerate(row):
                self.cellules[i][j].delete(0, tk.END)
                self.cellules[i][j].insert(0, value)
 

    def sauvegarder(self):
        """Sauvegarde les données des cellules dans un fichier CSV"""
        filetypes = (
            ('Fichiers CSV', '*.csv'),
            ('Tous les fichiers', '*.*')
        )

        # Demande à l'utilisateur de choisir un emplacement pour sauvegarder
        filename = fd.asksaveasfilename(defaultextension=".csv", filetypes=filetypes)

        if filename:
            try:
                with open(filename, 'w', encoding='utf-8', newline='') as f:
                    writer = csv.writer(f, delimiter=self.delimiteur)
                    for row in self.cellules:
                        writer.writerow([cell.get() for cell in row])

                messagebox.showinfo("Succès", "Fichier sauvegardé avec succès !")

            except PermissionError:
                print("Erreur : Permission refusée.")  # Debug
                messagebox.showerror("Erreur", "Impossible de sauvegarder le fichier : Permission refusée.")
            except Exception as e:
                print(f"Erreur inattendue : {e}")  # Debug
                messagebox.showerror("Erreur", f"Impossible de sauvegarder le fichier :\n{str(e)}")

    def update_buttons(self):
        """Met à jour l'état des boutons"""
        self.precedent.config(state=tk.NORMAL if self.index_actuel > 0 else tk.DISABLED)
        self.suivant.config(state=tk.NORMAL if self.index_actuel < len(self.history) - 1 else tk.DISABLED)

    def create_bande(self):
        """Crée une barre d'outils avec des boutons"""
        self.bande = ttk.Frame(self.root, padding=10, bootstyle="secondary")
        self.bande.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        # Boutons dans la barre d'outils
        self.precedent = ttk.Button(self.bande, text='←', command=self.précédent, state=DISABLED, bootstyle="primary")
        self.precedent.pack(side="left", padx=5, pady=5)

        self.suivant = ttk.Button(self.bande, text='→', command=self.suivant, state=DISABLED, bootstyle="primary")
        self.suivant.pack(side="left", padx=5, pady=5)

        ouvrir = ttk.Button(self.bande, text="Ouvrir", command=self.navigateur, bootstyle="success")
        ouvrir.pack(side="left", padx=5, pady=5)

        fusionner_csv = ttk.Button(self.bande, text="Fusionner CSV", command=self.fusionner_fichiers_csv, bootstyle="info")
        fusionner_csv.pack(side="left", padx=5, pady=5)

        fusionner_lignes = ttk.Button(self.bande, text="Fusionner Lignes", command=self.fusionner_lignes, bootstyle="info")
        fusionner_lignes.pack(side="left", padx=5, pady=5)

        sauvegarder = ttk.Button(self.bande, text="Sauvegarder", command=self.sauvegarder, bootstyle="warning")
        sauvegarder.pack(side="left", padx=5, pady=5)

        statistiques = ttk.Button(self.bande, text="Statistiques", command=self.calculer_statistiques, bootstyle="danger")
        statistiques.pack(side="left", padx=5, pady=5)

        # Menu déroulant pour changer le délimiteur
        delimiteur_menu = ttk.Menubutton(self.bande, text="Délimiteur", bootstyle="light")
        delimiteur_menu.menu = ttk.Menu(delimiteur_menu, tearoff=0)
        delimiteur_menu["menu"] = delimiteur_menu.menu

        delimiteur_menu.menu.add_command(label="Virgule (,)", command=lambda: self.changer_delimiteur(','))
        delimiteur_menu.menu.add_command(label="Point-virgule (;)", command=lambda: self.changer_delimiteur(';'))
        delimiteur_menu.menu.add_command(label="Tabulation (\\t)", command=lambda: self.changer_delimiteur('\t'))

        delimiteur_menu.pack(side="right", padx=5, pady=5)

        # Initialiser les boutons
        self.update_buttons()

    def calculer_statistiques(self):
        """Calcule et affiche des statistiques sur les données des cellules"""
        try:
            valeurs = []

            # Parcourir toutes les cellules pour récupérer les valeurs numériques
            for row in self.cellules:
                for cell in row:
                    contenu = cell.get()
                    if contenu.strip():  # Vérifie si la cellule n'est pas vide
                        try:
                            valeurs.append(float(contenu))  # Convertit en float si possible
                        except ValueError:
                            pass  # Ignore les valeurs non numériques

            if not valeurs:
                messagebox.showinfo("Statistiques", "Aucune donnée numérique trouvée.")
                return

            # Calcul des statistiques
            somme = sum(valeurs)
            moyenne = somme / len(valeurs)
            minimum = min(valeurs)
            maximum = max(valeurs)

            # Affichage des résultats
            messagebox.showinfo(
                " Vos Statistiques 😁",
                f"Somme : {somme}\nMoyenne : {moyenne}\nMinimum : {minimum}\nMaximum : {maximum}"
            )

        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite lors du calcul des statistiques :\n{str(e)}")

    def charger_csv_pandas(self):
        """Charge un fichier CSV dans les cellules en utilisant pandas pour permettre la manipulation et l’analyse de données"""
        filetypes = (
            ('Fichiers CSV', '*.csv'),
            ('Tous les fichiers', '*.*')
        )

        filename = fd.askopenfilename(filetypes=filetypes)

        if filename:
            try:
                print(f"Chargement du fichier : {filename}")  # Debug

                # Vérifier si le fichier est complètement vide
                with open(filename, 'r', encoding='utf-8') as f:
                    first_line = f.readline().strip()
                    if not first_line:
                        print("Le fichier est complètement vide.")  
                        messagebox.showerror("Erreur", "Le fichier CSV est complètement vide.")
                        return

                # Charger le fichier CSV avec pandas
                df = pd.read_csv(filename, delimiter=self.delimiteur)
                print(f"Fichier chargé avec succès :\n{df}")  

                # Vérifier si le fichier est vide après chargement
                if df.empty:
                    print("Le fichier est vide après chargement.")
                    messagebox.showerror("Erreur", "Le fichier CSV est vide.")
                    return

                # Vérifier si le fichier contient des colonnes
                if df.shape[1] == 0:
                    print("Le fichier ne contient aucune colonne.")  
                    messagebox.showerror("Erreur", "Le fichier CSV ne contient aucune colonne.")
                    return

                # Vérifier si le fichier contient des lignes
                if df.shape[0] == 0:
                    print("Le fichier ne contient aucune ligne.")  
                    messagebox.showerror("Erreur", "Le fichier CSV ne contient aucune ligne.")
                    return

                # Limiter les données au nombre de cellules disponibles
                max_rows, max_cols = len(self.cellules), len(self.cellules[0])
                for i in range(min(len(df), max_rows)):
                    for j in range(min(len(df.columns), max_cols)):
                        value = df.iloc[i, j]
                        if pd.isna(value):  # Vérifie les valeurs manquantes
                            value = ""
                        self.cellules[i][j].delete(0, tk.END)
                        self.cellules[i][j].insert(0, str(value))

                # Enregistrer l'état après chargement
                self.mise_a_jour()

            except pd.errors.EmptyDataError:
                print("Erreur : Le fichier est vide ou mal formaté.") 
                messagebox.showerror("Erreur", "Le fichier CSV est vide ou mal formaté.")
            except pd.errors.ParserError:
                print("Erreur : Problème d'analyse du fichier CSV.")  
                messagebox.showerror("Erreur", "Erreur lors de l'analyse du fichier CSV.")
            except UnicodeDecodeError:
                print("Erreur : Problème d'encodage.")  
                messagebox.showerror("Erreur", "Le fichier CSV contient des caractères non valides ou un encodage incorrect.")
            except Exception as e:
                print(f"Erreur inattendue : {e}") 
                messagebox.showerror("Erreur", f"Impossible de charger le fichier CSV :\n{str(e)}")


#Crée la fenêtre principale et lance l'application
if __name__ == "__main__":
    root = ttk.Window(themename="flatly")  
    app = MyCellule(root)
    root.mainloop()