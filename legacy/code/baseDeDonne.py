import mysql.connector
from db_manager import*
from tkinter import*
from tkinter import ttk
from datetime import datetime

class Gestion_base_de_donnee(FrameMultijoueur):
    def __init__(self, fenetre):
        super().__init__(fenetre)
        
        
        self.conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="",
            database="jeu_memoire"
        )
        
        self.cursor = self.conn.cursor()
        
       
    def creeTable(self):
        
        def Utilisateurs():
            def Envoyer():
                nom = entry_nom.get()
                age = entry_email.get()
                mail = entry_mdp.get()
                
                self.cursor.execute("INSERT INTO utilisateur (pseudo,email) VALUES(%s,%s)",(nom,mail))
                self.conn.commit()
                
                
                self.cursor.execute("SELECT * FROM utilisateur ")
                res = self.cursor.fetchall()
                liste = []
                for u in res:
                    if u not in liste:
                        liste.append(u)
                
                print(liste)
                for _ in liste:
                    self.tree.insert("","end",values=(_))
                        
                
                self.conn.commit()
            app = CTk()
            app.title("Formulaire Stylé")
            app.geometry("400x400")

            titre = CTkLabel(app, text="Formulaire de contact", font=("Arial", 22, "bold"))
            titre.pack(pady=20)

            entry_nom = CTkEntry(app, placeholder_text="Nom")
            entry_nom.pack(pady=10, ipadx=10, ipady=5)

            entry_email = CTkEntry(app, placeholder_text="Email")
            entry_email.pack(pady=10, ipadx=10, ipady=5)

            entry_mdp = CTkEntry(app, placeholder_text="Mot de passe", show="*")
            entry_mdp.pack(pady=10, ipadx=10, ipady=5)

            btn_envoyer = CTkButton(app, text="Envoyer" ,command=Envoyer)
            btn_envoyer.pack(pady=20)

            app.mainloop()    
            
        self.nomUtilisateur = CTkLabel(self.frameHaut,width=5,text="Toha" ,font=("",30),text_color="white")
        self.nomUtilisateur.place(x=610,y=70)
        self.Utilisateur = CTkButton(self.frameHaut,width=50,height=50,text="🧑",fg_color="blue",font=("",30))
        self.Utilisateur.place(x=630,y=10)
      
        self.Utilisateur.configure(command= Utilisateurs)
        requettes = [
       
        """
        CREATE TABLE IF NOT EXISTS utilisateur (
            id_utilisateur INT AUTO_INCREMENT PRIMARY KEY,
            pseudo VARCHAR(50) NOT NULL,
            email VARCHAR(100)
        );
            """,
        """
            CREATE TABLE IF NOT EXISTS Partie (
            id_partie INT AUTO_INCREMENT PRIMARY KEY,
            Nom_joueur VARCHAR(30),
            date_heure DATETIME DEFAULT CURRENT_TIMESTAMP,
            mode ENUM('solo', 'multijoueur', 'vs_ordi') NOT NULL
        );
        """,
        """
            CREATE TABLE IF NOT EXISTS Score (
                id_score INT AUTO_INCREMENT PRIMARY KEY,
                id_utilisateur INT,
                id_partie INT,
                score_obtenu INT,
                temps_ecoule TIME,
                erreur INT,
                coups INT,
                FOREIGN KEY(id_utilisateur) REFERENCES utilisateur(id_utilisateur),
                FOREIGN KEY(id_partie) REFERENCES Partie(id_partie)
        );
        """]
        self.conn.commit()
        
        for requette in requettes:
            self.cursor.execute(requette)
        self.conn.commit()
        
        
        
        
       
            
    def Afficher(self):
        
        tree2 = ttk.Treeview(self.t,columns=("id","nom_joueur","mode","heure"), show="headings")
        tree2.heading("id",text="partie-jouer")
        tree2.heading("nom_joueur",text="Nom joueur")
        tree2.heading("mode",text="mode")
        tree2.heading("heure",text="jour-date-heure")
        
        
        tree2.column("id",anchor="center")
        tree2.column("nom_joueur",anchor="center")
        tree2.column("mode",anchor="center")
        tree2.column("heure",anchor="center")
        tree2.pack()
        
        self.tree = ttk.Treeview(self.t,columns=("identifient","nom","pseudo","email"), show="headings",height=10)
        self.tree.heading("identifient",text="joueurs")
        self. tree.heading("nom",text="Nom")
        self.tree.heading("pseudo",text="pseudo")
        self.tree.heading("email",text="email")
        self.tree.pack(pady=0, padx=0, fill="both", expand=True)
        self.tree.pack()
        
        self.cursor.execute("SELECT * FROM utilisateur ")
        res = self.cursor.fetchall()
        
        for i in res:
            self.tree.insert("","end",values=(i))
            print(i[0],i[3],i[2],i[1])

        
        
        self.cursor.execute('SELECT * FROM Partie')
        resue = self.cursor.fetchall()
        for i in resue:
            tree2.insert("","end",values=(i[0],i[3],i[2],i[1]))
            print("c'est i",i[0],i[3],i[2],i[1])
        
        self.cursor.execute('SELECT * FROM Score')
        resultat = self.cursor.fetchall()
        for i in resultat:
            print(i)
            self.tree3.insert("","end",values=(i[3],i[4],i[5],i[6]))

    
    
        
    
            
    
            
if __name__=='__main__':
    jeu = CTk()
    #fenetre = Memory_game(jeu)
    fenetre = Gestion_base_de_donnee(jeu)
    fenetre.frameMultijoueur()
    fenetre.frame1()
    fenetre.frame2()
    fenetre.frame3()
    fenetre.Afficher()
    fenetre.creeTable()
    fenetre.Musique()
    #fenetre.insertion_des_donnee()
    jeu.resizable(width=False,height=False)
    fenetre.run()