from memory_game2 import*


class FrameMultijoueur(Memory_game2):
    def __init__(self, fenetre):
        super().__init__(fenetre)
        self.joueur1 = True
        self.joueur2 = False
        self.j1s = 0
        self.j2s =0
        shuffle(self.listeImages)
        self.listeImage = []
      
    def frameMultijoueur(self):
        self.frMultijoueur = CTkFrame(self.fenetre)
        self.frMultijoueur.place(relwidth=1,relheight=1)
        

        
        self.frameMultijoueurConteneur =  CTkFrame(self.frMultijoueur,width=750,height=750+50)
        self.frameMultijoueurConteneur.pack()
        
        self.frMultijoueurOrdi = CTkFrame(self.fenetre,fg_color="blue")
        self.frMultijoueurOrdi.place(relwidth=1,relheight=1)
        
        self.frMultijoueurSolo = CTkFrame(self.fenetre,fg_color="blue")
        self.frMultijoueurSolo.place(relwidth=1,relheight=1)
        
        self.frameMutijoueurChoixBleu = CTkFrame(self.frameMultijoueurConteneur,width=750,height=750,fg_color="blue")
        self.frameMutijoueurChoixBleu.pack_propagate(False)
        self.frameMutijoueurChoixBleu.place(relwidth=1,relheight=0.5)
        
        
        self.frameMutijoueurChoixRouge = CTkFrame(self.frameMultijoueurConteneur,width=750,height=750,fg_color="red")
        self.frameMutijoueurChoixRouge.pack_propagate(False)
        self.frameMutijoueurChoixRouge.place(relwidth=1,relheight=0.5, y=380)
        
        
        self.ButtonMultijAmi= CTkButton(self.frameMutijoueurChoixBleu,text="Jouer à 2",font=("Comic Sans MS",30),width=200,height=100,command=lambda:self.changerFrame(self.frMultijoueurSolo))
        self.ButtonMultijAmi.pack(expand=True)
        
        
        self.ButtonMultijPC = CTkButton(self.frameMutijoueurChoixRouge,text="Jouer contre \nl’ordinateur",font=("Comic Sans MS",30),width=200,height=100,command=lambda:self.changerFrame(self.frMultijoueurOrdi))
        self.ButtonMultijPC.pack(expand=True)
        
        
        self.buttonRetoure = CTkButton(self.frameMultijoueurConteneur,text="⬅️",width=50,height=70,font=("",50),text_color="white",fg_color="green",command=lambda:self.changerFrame(self.fr1))
        self.buttonRetoure.place(x=0,y=0)
        
        self.buttonRetoure = CTkButton(self.frMultijoueurSolo,text="⬅️",width=50,height=70,font=("",50),text_color="white",fg_color="green",command=lambda:self.changerFrame( self.frMultijoueur))
        self.buttonRetoure.place(x=700,y=0) 
        
        self.GrilleMultijoueur()
        self.creegrgrille2()
        
    def GrilleMultijoueur(self):
        self.frameConteneur = CTkFrame(self.frMultijoueurSolo)
        self.frameConteneur.pack(expand=True)
        labelJ1 = Label(self.frameConteneur,text="J1: 0",font=("",20),bg="red")
        labelJ1.grid(row=5,column=1)
        labelJ2 = Label(self.frameConteneur,text="J2: 0",font=("",20))
        labelJ2.grid(row=5,column=2)
        buttons = []
        for i in range(4):
            l = []
            for j in range(4):
                button = Button(self.frameConteneur,image=self.photo ,text = "?",width=self.largeur,height=self.largeur,command=lambda i=i,j=j:retournement(i,j),font=("",50),fg="red")
                button.grid(row=i,column=j)
                l.append(button)
            buttons.append(l)
        
        liste = ['A','B','C','D','E','F','G','H']*2
        shuffle(liste)
        carteretourner = []
        def retournement(lig,col):
            #global joueur1 ,joueur2,j1s,j2s
            index = lig*4+col
            buttons[lig][col].configure(text=liste[index])
            carteretourner.append((lig,col))
            if len(carteretourner)==2:
                l1,c1 = carteretourner[0]
                l2,c2 = carteretourner[1]
                if liste[l1*4+c1] == liste[l2*4+c2]:
                    print("j1",self.joueur1)
                    print("j2",self.joueur2)
                    if self.joueur1 == True:
                        self.j1s+=1
                        labelJ1.config(text=f"J1: {self.j1s}")
                    else:
                        m=[]
                        self.j2s+=1
                        labelJ2.config(text=f"J2: {self.j2s}") 
                        
                    print("carte trouver")
                else:
                    self.fenetre.after(500,lambda:refermer(l1,c1,l2,c2))
                    if self.joueur1 ==True :
                        self.joueur1 = False
                        self.joueur2 = True
                        labelJ2.config(bg="red")
                        labelJ1.config(bg="white")
                        
                    else:
                        self.joueur2 = False
                        self.joueur1 = True
                        labelJ1.config(bg="red")
                        labelJ2.config(bg="white")
                        
                    
                    print(self.joueur1)
                print(2)
                carteretourner.clear()
                
        def refermer(l1,c1,l2,c2):
            buttons[l1][c1].config(text="?")
            buttons[l2][c2].config(text="?")
                
    def creegrgrille2(self):

       
        self.frameOrdi = CTkFrame(self.frMultijoueurOrdi)
        self.frameOrdi.pack(expand=True)
        
        labelJ1 = Label(self.frameOrdi, text="J1: 0", font=("", 20), bg="red")
        labelJ1.grid(row=5, column=1)
        labelJ2 = Label(self.frameOrdi, text="Ordi: 0", font=("", 20))
        labelJ2.grid(row=5, column=2)
        

        buttons = []
        for i in range(4):
            l = []
            for j in range(4):
                button = Button(self.frameOrdi,image=self.photo, text="?",width=self.largeur,height=self.largeur, command=lambda i=i, j=j: retournement(i, j), font=("", 50), fg="red")
                button.grid(row=i, column=j)
                l.append(button)
            buttons.append(l)

        self.joueur1 = True  # True = Joueur humain, False = Ordinateur
        self.j1s = 0
        self.j2s = 0
        liste = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'] * 2
        shuffle(liste)  # Mélange les cartes au début
        carteretourner = []


        def retournement(lig, col):
            index = lig * 4 + col

            if buttons[lig][col]['text'] != "?":  # Empêcher de cliquer sur une carte déjà retournée
                return

            buttons[lig][col].configure(text=liste[index] ,image = self.images[index])
            carteretourner.append((lig, col))

            if len(carteretourner) == 2:
                l1, c1 = carteretourner[0]
                l2, c2 = carteretourner[1]

                if liste[l1 * 4 + c1] == liste[l2 * 4 + c2]:  # Si c'est une paire
                    if self.joueur1:
                        self.j1s += 1
                        labelJ1.config(text=f"J1: {self.j1s}")
                    else:
                        self.j2s += 1
                        labelJ2.config(text=f"Ordi: {self.j2s}")
                        self.fenetre.after(1000, tour_ordinateur) 
                        
                else:
                    self.fenetre.after(500, lambda: refermer(l1, c1, l2, c2))
                    joueur_suivant()

                carteretourner.clear()


        def refermer(l1, c1, l2, c2):
        
            buttons[l1][c1].config(text="?")
            buttons[l2][c2].config(text="?")


        def joueur_suivant():
            self.joueur1 = not self.joueur1

            if self.joueur1:
                labelJ1.config(bg="red")
                labelJ2.config(bg="white")
            else:
                labelJ1.config(bg="white")
                labelJ2.config(bg="red")
                self.fenetre.after(1000, tour_ordinateur)

        n = []
        def tour_ordinateur():
            
            coups_possibles = [(i, j) for i in range(4) for j in range(4) if buttons[i][j]['text'] == "?"]

            if len(coups_possibles) < 2:
                return  

            coup1 = choice(coups_possibles)
            retournement(*coup1)
            n.append(coup1)
            
            coups_possibles.remove(coup1)  # Empêcher l'ordinateur de retourner la même carte deux fois
            coup2 = choice(coups_possibles)
            n.append(coup2)
            self.fenetre.after(1000, lambda: retournement(*coup2))  # Retourne la deuxième carte après 1s


       

if __name__=='__main__':
    jeu = CTk()
    #fenetre = Memory_game(jeu)
    fenetre = FrameMultijoueur(jeu)
    fenetre.frameMultijoueur()
    fenetre.frame1()
    fenetre.frame2()
    fenetre.frame3()
    
    jeu.resizable(width=False,height=False)
    fenetre.run()