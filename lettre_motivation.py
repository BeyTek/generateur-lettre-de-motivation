import tkinter as tk
from tkinter import filedialog, messagebox,ttk
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO
from ttkthemes import ThemedStyle

class LettreMotivationGenerator:
    def __init__(self, root):
        self.root = root
        root.geometry("600x600")
        self.root.title("Générateur de Lettre de Motivation")
        style = ThemedStyle(root)
        style.set_theme("plastik")
        self.nom_var = tk.StringVar()
        self.prenom_var = tk.StringVar()
        self.nom_entreprise_var = tk.StringVar()
        self.poste_var = tk.StringVar()
        self.adresse_var = tk.StringVar()
        self.code_postal_ville_var = tk.StringVar()
        self.telephone_var = tk.StringVar()
        self.email_var = tk.StringVar()

        frame_nom_exp = tk.Frame(root)
        frame_nom_exp.pack()
        
        tk.Label(frame_nom_exp, text="Nom:").grid(row=0, column=0)
        self.nom_entry = tk.Entry(frame_nom_exp, textvariable=self.nom_var)
        self.nom_entry.grid(row=0, column=1)

        tk.Label(frame_nom_exp, text="Prénom:").grid(row=1, column=0)
        self.prenom_entry = tk.Entry(frame_nom_exp, textvariable=self.prenom_var)
        self.prenom_entry.grid(row=1, column=1)

        tk.Label(frame_nom_exp, text="Nom de l'entreprise:").grid(row=2, column=0)
        self.nom_entreprise_entry = tk.Entry(frame_nom_exp, textvariable=self.nom_entreprise_var)
        self.nom_entreprise_entry.grid(row=2, column=1)

        tk.Label(frame_nom_exp, text="Nom du poste:").grid(row=3, column=0)
        self.poste_entry = tk.Entry(frame_nom_exp, textvariable=self.poste_var)
        self.poste_entry.grid(row=3, column=1)

        tk.Label(root, text="Adresse:").pack()
        self.adresse_entry = tk.Entry(root, textvariable=self.adresse_var)
        self.adresse_entry.pack()

        tk.Label(root, text="Code postal, Ville:").pack()
        self.code_postal_ville_entry = tk.Entry(root, textvariable=self.code_postal_ville_var)
        self.code_postal_ville_entry.pack()

        tk.Label(root, text="Téléphone:").pack()
        self.telephone_entry = tk.Entry(root, textvariable=self.telephone_var)
        self.telephone_entry.pack()

        tk.Label(root, text="Adresse e-mail:").pack()
        self.email_entry = tk.Entry(root, textvariable=self.email_var)
        self.email_entry.pack()

        ttk.Button(root, text="Générer Lettre", command=self.generer_lettre).pack()
        ttk.Button(root, text="Convertir en PDF", command=self.convertir_en_pdf).pack()

    def generer_lettre(self):
        nom = self.nom_var.get()
        prenom = self.prenom_var.get()
        nom_entreprise = self.nom_entreprise_var.get()
        poste = self.poste_var.get()
        adresse = self.adresse_var.get()
        code_postal_ville = self.code_postal_ville_var.get()
        telephone = self.telephone_var.get()
        email = self.email_var.get()

        # Obtenir la date actuelle au format jj/mm/aaaa
        date_actuelle = datetime.now().strftime("%d/%m/%Y")

        lettre = f"""{prenom} {nom}
        {adresse}
        {code_postal_ville}
        {telephone}
        {email}

        Fait le {date_actuelle} a {code_postal_ville}

        Objet : Candidature au poste de {poste}

        Madame, Monsieur,
        Je me permets de vous adresser ma candidature pour le poste de {poste} au sein de votre entreprise {nom_entreprise}.En tant que personne dévouée et motivée, je suis convaincue que mon profil correspond aux compétences recherchées pour ce poste.
        Intégrer votre équipe serait une opportunité fantastique pour moi de contribuer au succès de {nom_entreprise} tout en développant mes compétences professionnelles dans ce domaine.
        Je me tiens à votre disposition pour discuter de ma candidature en détail et pour répondre à toutes les questions que vous pourriez avoir. Je vous remercie de l'attention que vous porterez à ma candidature et j'espère avoir l'occasion de contribuer au succès continu de {nom_entreprise}.
        Cordialement,
        {prenom} {nom}
        """

        nom_du_fichier = f'LettreDeMotivation_{prenom}_{nom_entreprise}.txt'

        with open(nom_du_fichier, 'w') as fichier:
                    fichier.write(lettre)

        messagebox.showinfo("Lettre Générée", f"La lettre de motivation a été générée avec succès dans le fichier {nom_du_fichier}.")

    def convertir_en_pdf(self):
        # Demander à l'utilisateur de sélectionner le fichier texte à convertir en PDF
        fichier_texte = filedialog.askopenfilename(filetypes=[("Fichiers texte", "*.txt")])

        if fichier_texte:
            # Lire le contenu du fichier texte
            with open(fichier_texte, "r") as fichier_txt:
                contenu = fichier_txt.readlines()

            nom = self.nom_var.get()
            prenom = self.prenom_var.get()
            nom_entreprise = self.nom_entreprise_var.get()
            poste = self.poste_var.get()
            adresse = self.adresse_var.get()
            code_postal_ville = self.code_postal_ville_var.get()
            telephone = self.telephone_var.get()
            email = self.email_var.get()

            # Obtenir la date actuelle au format jj/mm/aaaa
            date_actuelle = datetime.now().strftime("%d/%m/%Y")

            # Créer un fichier PDF et écrire le contenu du fichier texte dedans
            nom_du_fichier_pdf = f'LettreMotivation_{prenom}_{nom_entreprise}.pdf'

            doc = SimpleDocTemplate(nom_du_fichier_pdf, pagesize=letter)
            styles = getSampleStyleSheet()
            styles["Normal"].fontSize += 4
            flowables = []

            for ligne in contenu:
            # Ajouter un espace entre "fait le" et "Objet" ainsi qu'entre "Objet" et "Madame, Monsieur,"
                if "fait le" or "objet" in ligne:
                    flowables.append(Spacer(1, 12))  # Ajouter un espace de 12 points
                p = Paragraph(ligne.strip(), styles["Normal"])
                flowables.append(p)

            doc.build(flowables)

            messagebox.showinfo("PDF Généré", f"Le PDF a été créé avec succès dans le fichier {nom_du_fichier_pdf}.")


         
if __name__ == "__main__":
    root = tk.Tk()
    app = LettreMotivationGenerator(root)
    root.mainloop()