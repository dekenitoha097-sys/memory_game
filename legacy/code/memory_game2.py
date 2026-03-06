from h import*
import mysql.connector
from tkinter import ttk
from datetime import datetime

    

    
class Memory_game2(Memory_game):
    def __init__(self, fenetre):
        super().__init__(fenetre)
                    
        self.conn = mysql.connector.connect(
            host="localhost",
            user="root",
            password="",
            database="jeu_memoire"
        )
        self.cursor = self.conn.cursor()
        
        
    def changerTaillegrille(self, ligne, colonne):
        """Met à jour la taille de la grille et la recrée."""
        self.ligne = ligne
        self.colonne = colonne
        
        if self.ligne ==4 and self.colonne==3:
            self.largeur =230
            self.longeur = 160
            self.imagelongeur =230
            self.imagelargeur = 160 
            self.j = 3
            self.carteTrouver =0
            shuffle(self.listeImages)
            self.listeImage = self.listeImages[:6]*2
            print(self.listeImage)
            shuffle(self.listeImage)
            self.images.clear()
            self.im = [Image.open(img) for img in self.listeImage]
            self.images =[]
            for i in self.im:
                l = i.resize((self.imagelongeur,self.imagelargeur))
                self.images.append(l)
            self.images = [ImageTk.PhotoImage(img) for img in self.images]
            #for i in self.im:
            # l = i.resize((self.imagelongeur,self.imagelargeur))
                #self.images.append(ImageTk.PhotoImage(l))
                
        elif self.ligne ==4 and self.colonne==4:
            self.largeur =150
            self.longeur = 150
            self.imagelongeur =150
            self.imagelargeur = 150 
            self.carteTrouver =0
            ##self.listeImage = self.listeImage[::8]*2###@
            self.j = 4
            shuffle(self.listeImages)
            self.listeImage = self.listeImages[:8]*2
            shuffle(self.listeImage)
            self.images.clear()
            self.im = [Image.open(img) for img in self.listeImage]
            self.images =[]
            for i in self.im:
                l = i.resize((self.imagelongeur,self.imagelargeur))
                self.images.append(l)
            self.images = [ImageTk.PhotoImage(img) for img in self.images]
            #for i in self.im:
                #l = i.resize((self.imagelongeur,self.imagelargeur))
                #self.images.append(ImageTk.PhotoImage(l)) 

        elif self.ligne ==5 and self.colonne==4:
            self.largeur = 150
            self.longeur = 125
            self.imagelongeur =150
            self.carteTrouver =0
            self.imagelargeur = 125 
            self.j = 4
            shuffle(self.listeImages)
            self.listeImage = self.listeImages[:10]*2
            shuffle(self.listeImage)
            self.images.clear()
            self.im = [Image.open(img) for img in self.listeImage]
            self.images =[]
            for i in self.im:
                l = i.resize((self.imagelongeur,self.imagelargeur))
                self.images.append(l)
            self.images = [ImageTk.PhotoImage(img) for img in self.images]
            
            #for i in self.im:
                #l = i.resize((self.imagelongeur,self.imagelargeur))
                #self.images.append(ImageTk.PhotoImage(l))  
                
            
            print(self.listeImage)

        elif self.ligne ==6 and self.colonne==5:
            self.largeur = 140
            self.longeur =100
            self.carteTrouver =0
            self.imagelongeur =140
            self.imagelargeur = 100
            self.j = 5
            shuffle(self.listeImages)
            self.listeImage = self.listeImages[:15]*2
            self.images.clear()
            shuffle(self.listeImage)
            self.images.clear()
            self.im = [Image.open(img) for img in self.listeImage]
            self.images =[]
            for i in self.im:
                l = i.resize((self.imagelongeur,self.imagelargeur))
                self.images.append(l)
            self.images = [ImageTk.PhotoImage(img) for img in self.images]
            
            #or i in self.im:
            #   l = i.resize((self.imagelongeur,self.imagelargeur))
            #    self.images.append(ImageTk.PhotoImage(l)) 
            

        print(f"Nouvelle taille de la grille : {self.ligne}x{self.colonne}")
        self.CreeGrille()   

    def retournementCarte(self,lig,col):
        
        index = lig*self.j+col
        
        self.n[lig][col].config(image = self.images[index])
       
           
        if (lig,col) not in self.carteRetourner :
            self.carteRetourner.append((lig,col))
        else:
            pass
        if len(self.carteRetourner)==2:
            self.coup+=1
            self.couplabel.configure(text=f"COUPS: {self.coup}")
            l1,c1 = self.carteRetourner[0]
            l2,c2 = self.carteRetourner[1]
            if self.listeImage[l1*self.j+c1] == self.listeImage[l2*self.j+c2]:
                self.n[l1][c1].config( state=DISABLED)
                self.n[l2][c2].config( state=DISABLED)
                self.carteTrouver +=1
                print(len(self.listeImage))
                print(self.carteTrouver)
                
                
                if self.carteTrouver == len(self.listeImage)//2:
                    
                    self.cursor.execute("INSERT INTO Score (score_obtenu ,temps_ecoule,erreur,coups) VALUES(%s,%s,%s,%s)", (self.score,  self.second, self.erreur, self.coup))
                    print("c'est bien fait")
                    self.conn.commit()
                    self.cursor.execute("INSERT INTO Partie (Nom_joueur,date_heure) VALUES(%s,%s)",(self.nomJouer,datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                    self.conn.commit()
                    
                 
                    
                    print("felicitation")
                    def gagner():
                        self.T = False
                        pauseFrame = CTkFrame(self.CntButton,width=500,height=500,fg_color="blue",bg_color="yellow")
                        pauseFrame.pack_propagate(False)
                        pauseFrame.place(relx=0.5, rely=0.5, anchor="center")
                        textGagner = CTkLabel(pauseFrame,text ="Félicitations !\n Vous avez gagné ! 🏆",font=("",50))
                        textGagner.place(relx=0.5,rely=0.17, anchor="center")
                        frame = CTkFrame(pauseFrame,fg_color="green",width=300)
                        frame.pack(expand=True)
                        def etoile1():
                            labelEtoile2 = CTkLabel(frame,text="⭐",text_color="orange",font=("",50))
                            labelEtoile2.place(relx=0.3, rely=0.5, anchor="center")
                        def etoile2():
                            labelEtoile = CTkLabel(frame,text="⭐",text_color="orange",font=("",70))
                            labelEtoile.place(relx=0.5, rely=0.3, anchor="center")
                        def etoile3():
                            labelEtoile1 = CTkLabel(frame,text="⭐",text_color="orange",font=("",50))
                            labelEtoile1.place(relx=0.7, rely=0.5, anchor="center")
                            
                        def dg(ligne,colonne,s1,s2):
                            if self.ligne ==ligne and self.colonne ==colonne:
                                if self.second <=s1:
                                    self.fenetre.after(1000,lambda:etoile1())
                                    self.fenetre.after(2000,lambda:etoile2())
                                    self.fenetre.after(3000,lambda:etoile3())
                                elif self.second >s1 and self.second<s2:
                                    self.fenetre.after(1000,lambda:etoile1())
                                    self.fenetre.after(2000,lambda:etoile2())
                                else:
                                    self.fenetre.after(1000,lambda:etoile1())
                        dg(4,3,25,50)
                        dg(4,4,35,60)
                        dg(5,4,45,80)
                        dg(6,5,80,100)
                        def reini():
                            self.T =True
                            self.second = 0
                            self.templabel.configure(text = f"00 : 00 : 00")
                            for i in self.n:
                                for j in i:
                                    j.configure(image=self.photo)
                            pauseFrame.place(relwidth=0,relheight=0)
                            self.carteTrouver= 0
                            shuffle(self.listeImage)
                            self.images.clear()
                            self.im = [Image.open(img) for img in self.listeImage]
                            self.images =[]
                            for i in self.im:
                                l = i.resize((self.imagelongeur,self.imagelargeur))
                                self.images.append(l)
                            self.images = [ImageTk.PhotoImage(img) for img in self.images]\
                                
                            for i in self.n:
                                for j in i:
                                    j.config(state="active")
                        buttRed = CTkButton(pauseFrame,text="🔄",text_color="white",font=("",50),width=50,fg_color="red",command=reini)
                        buttRed.place(relx=0.7, rely=0.7, anchor="center")
                        buttRed2 = CTkButton(pauseFrame,text="🏠",text_color="white",font=("",50),width=50,fg_color="red",border_color="yellow",command=lambda:self.changerFrame(self.fr1))
                        buttRed2.place(relx=0.3, rely=0.7, anchor="center")
                    self.fenetre.after(1000,lambda:gagner())
                self.score+=1
                self.scorelabel.configure(text=f"SCORE: {self.score}")
                print("wow")
            else:
                self.fenetre.after(500,lambda:self.reinitialiserCarte(l1,c1,l2,c2))
                self.erreur+=1
                self.erreurlabel.configure(text=f"ERREUR: {self.erreur}")
            self.carteRetourner.clear()
           
            
    def reinitialiserCarte(self,l1,c1,l2,c2):
        self.n[l1][c1].config(image=self.photo)
        self.n[l2][c2].config(image=self.photo)
    
   