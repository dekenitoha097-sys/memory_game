from customtkinter import*
from tkinter import*
from CTkMessagebox import CTkMessagebox
from PIL import Image ,ImageTk
from random import*
import pygame
from tkinter import ttk




class Memory_game:
    def __init__(self,fenetre):
        self.fenetre = fenetre
        self.fenetre.geometry("800x800")
        self.fenetre.title("Jeu Memoire")
        self.ligne = 4
        self.colonne = 4
        self.largeur = 150
        self.longeur = 150
        self.carteRetourner = []
        self.j = 4
        self.imagelongeur = 150
        self.imagelargeur = 150
        self.coup = 0
        self.erreur = 0
        self.score = 0
        self.second = 0
        self.minute = 0
        self.T = False
        self.carteTrouver = 0
        self.nomJouer = "dekeni Toha"
        
        
        imag = 'bleu.jpeg'
        image = Image.open(imag)
        image = image.resize((300,300))
        self.photo = ImageTk.PhotoImage(image)
        
        self.listeImagechoix = ["voiture.jpeg","panda.jpeg","dessin.jpeg","nature.jpeg","manga.jpeg","sport.jpeg"
                                ,"film.jpeg","fruit.jpeg","cartejeu.jpeg","afrique.jpeg","drapeau.jpeg","musique.jpeg",
                                "touristique.jpeg","criquet.jpeg","science.jpeg","jeuvideo.jpeg","planete.jpeg"
                                ,"maison.jpeg","ville.jpeg","acteurs.jpeg"]*2
        
        print("c'est",len(self.listeImagechoix))
        
        print(len(self.listeImagechoix))
        
        ima = 'Spider.jpeg'
        im = [Image.open("C:/Users/deken/Desktop/PP/image_voiture/"+img) for img in self.listeImagechoix]
        self.imageschoix =[]
        for i in im:
            c = i.resize((100,100))
            self.imageschoix.append(c)
        self.imageschoix = [ImageTk.PhotoImage(img) for img in self.imageschoix]
        
        #i = i.resize((100,100))
        #self.phot = ImageTk.PhotoImage(i)
        
        self.listeImages = ['grenouille.jpeg','cerf.jpeg','pigeon.jpeg','poisson.jpeg',
                           'tortue.jpeg','serpent.jpeg','spider.jpeg','tigre.jpeg'
                           ,'Camel.jpeg','zebre.jpeg','panda.jpeg','kangaroo.jpeg','elephant.jpeg'
                           ,'Dog.jpeg','escargot.jpeg','sheep.jpeg']
        self.liste2 = ['voiture1.jpeg','voiture2.jpeg','voiture3.jpeg','voiture4.jpeg']*10
        
        shuffle(self.listeImages)
        self.listeImage = []
        
        self.im = [Image.open(img) for img in self.listeImage]
        self.images =[]
        for i in self.im:
            l = i.resize((self.imagelongeur,self.imagelargeur))
            self.images.append(l)
        self.images = [ImageTk.PhotoImage(img) for img in self.images]
        #logos = [PhotoImage(file=img.strip()) for img in liste]
    def frame1(self):
        self.fr1 = CTkFrame(self.fenetre)
        self.fr1.place(relwidth=1,relheight=1) 
        self.fr2 = CTkFrame(self.fenetre)
        self.fr2.place(relwidth=1,relheight=1)
        self.fr3 = CTkFrame(self.fenetre)
        self.fr3.place(relwidth=1,relheight=1)
        
        self.tab1 = CTkTabview(self.fr1,width=750,height=750,fg_color="#FCC6FF")
        self.tab1.pack()
        
        self.tab = self.tab1.add("ACCUEIL")
        self.ta = self.tab1.add("PARAMÈTRE")
        self.t = self.tab1.add("PROFILE")
        
                
        style = ttk.Style()
        style.theme_use("default") 

       
        style.configure("Treeview",
            background="#2e2e2e",
            foreground="white",
            rowheight=25,
            fieldbackground="#2e2e2e",
            font=("Helvetica", 10)
        )

       
        style.configure("Treeview.Heading",
            background="#444",
            foreground="cyan",
            font=("Helvetica", 11, "bold")
        )

        
        style.map("Treeview",
            background=[('selected', '#6A9FB5')],
            foreground=[('selected', 'red')]
)
        
        self.tree3 = ttk.Treeview(self.t,columns=("score","temps","erreur","coups"), show="headings")
        self.tree3.heading("score",text="Score")
        self.tree3.heading("temps",text="Temps")
        self.tree3.heading("erreur",text="Erreur")
        self.tree3.heading("coups",text="nombre coups")
        
        self.tree3.column("score",anchor="center")
        self.tree3.column("temps",anchor="center")
        self.tree3.column("erreur",anchor="center")
        self.tree3.column("coups",anchor="center")
        self.tree3.pack()      
        
        for ta in self.tab1._segmented_button._buttons_dict.values():
            ta.configure(font=("Arial", 20, "bold"))
        self.frameHaut = CTkFrame(self.tab,fg_color="#4B164C",height=100,width=750)
        self.frameHaut.pack_propagate(False)
        
        self.frameHaut.pack(fill="both",expand=False)
        self.frameTitre = CTkFrame(self.frameHaut,fg_color="#4B164C")
        self.frameTitre.pack(expand=True)
       
        ls = ['M','E','M','O','R','Y','🎲']
        ls2 = ['C','H','A','L','L','E','N','G','E']
        m = []
        m2 = []
        for j in range(7):
            self.lM = Label(self.frameTitre,text=ls[j],font=("Arial",20,"bold"))
            if ls[j]=='M':
                self.lM.config(bg="red")
            self.lM.grid(row=0,column=j,padx=5,pady=5)
            m.append(self.lM)
            
        for i in range(9):
            self.lM = Label(self.frameTitre,text=ls2[i],font=("Arial",20,"bold"))
            self.lM.grid(row=1,column=i+1,padx=5,pady=5)
            m2.append(self.lM)

        
        self.framebas = CTkFrame(self.tab,fg_color="#EDF4C2",height=50,width=750)
        
        
        
        
        
        
        import random
        self.index = 0

        def t():
           
            liste = [ "📌 Amusez-vous et testez votre mémoire !",
                    "🧠 Seuls les meilleurs mémorisent tout... Serez-vous à la hauteur ?",
                    "🎮 Préparez-vous à un défi cérébral !",
                    "✨ Un jeu... Un défi... Une aventure pour votre mémoire !",
                    "😂 Si vous oubliez ce jeu, c'est que vous devez y jouer encore plus !"
                    ]
            self.index = random.randrange(4)
            s = liste[self.index]
            self.labelBas.configure(text=s,font=("Arial",20,"bold"))
            self.labelBas2.configure(text=s,font=("Arial",20,"bold"))
            self.labelBas3.configure(text=s,font=("Arial",20,"bold"))
            self.fenetre.after(3000,t)
            

        self.fenetre.after(1000,t)
        self.labelBas = CTkLabel(self.framebas,text="",width=750)
        self.labelBas.pack()
        
        
        
        self.conteiner = CTkFrame(self.tab)
        self.conteiner.pack(expand=True)
        
        self.buttonL = CTkButton(self.conteiner,width=200,height=100,text="  JEU LIBRE 📌 ",font=("Comic Sans MS",30),fg_color="#A6CDC6",command=lambda:self.changerFrame(self.fr2))
        self.buttonL.grid(row=0,column=0)
        self.buttonL1 = CTkButton(self.conteiner,width=200,height=100,text="MODE DEFI 🧠",font=("Comic Sans MS",30),fg_color="#2DAA9E")
        self.buttonL1.grid(row=0,column=1)
        self.buttonL2 = CTkButton(self.conteiner,width=200,height=100,text="JEU A DEUX 🎮",font=("Comic Sans MS",30),fg_color="#4B164C" ,command=lambda:self.changerFrame(self.frMultijoueur))
        self.buttonL2.grid(row=1,column=0)
        self.buttonL3 = CTkButton(self.conteiner,width=200,height=100,text=" HISTOIRE 🎮 ",font=("Comic Sans MS",30),fg_color="#4B164C")
        self.buttonL3.grid(row=1,column=1)
        
        self.LabelPara =  CTkLabel(self.tab,text="Bienvenue dans Memory Challenge ! \nTestez votre mémoire et votre rapidité en retrouvant les paires le plus vite possible.\n Que vous soyez un joueur occasionnel ou un expert en concentration,\n ce jeu mettra vos capacités à l’épreuve avec différents niveaux de difficulté.\n À vous de jouer et de battre votre propre record ! 🎮💡\n",
                                   font=("",12,"italic"),fg_color="lightblue",width=750)
        self.LabelPara.pack()
        self.framebas.pack()
        
        
        imCarte = Image.open("cartes.png")
        imCarte = imCarte.resize((100,100))
        imcart = ImageTk.PhotoImage(imCarte)
        self.cartes = CTkLabel(self.frameHaut,image=imcart,text="")
        self.cartes.place(x=0,y=0)
        self.changerFrame(self.fr1)
        
        self.frameParametre = CTkFrame(self.ta,fg_color="red")
        self.frameParametre.place(relwidth=1,relheight=0.5)
        self.frameParametreApparence0 = CTkFrame(self.frameParametre,fg_color="blue")
        self.frameParametreApparence0.place(relwidth=1,relheight=0.5)
        self.texteApparence = CTkLabel(self.frameParametreApparence0,text="APPARENCE",font=("Arial",100))
        self.texteApparence.pack(expand=True)
        self.frameParametreApparence = CTkFrame(self.frameParametre,fg_color="green")
        self.frameParametreApparence.place(relwidth=1,relheight=0.5,y=173)
        
        
        for i in range(2):
            for j in range(6):
                button = CTkButton(self.frameParametreApparence,height=50,width=83)
                button.grid(row=i,column=j,pady=20,padx=20) 
           
        self.frameParametre = CTkFrame(self.ta,fg_color="orange")
        self.frameParametre.place(relwidth=1,relheight=0.5,y=350) 
        self.frameParametrePara = CTkFrame(self.frameParametre,fg_color="gray")
        self.frameParametrePara.place(relheight=1,relwidth=0.5) 
        listePara = ["CARTES NUMÉROTÉES: ","CACHER LES PAIRES: ","CACHER LE CHRONO: ","CACHER LES ERREURS: ","CACHER LES COUPS: "] 
        listebutt = ["ON","OFF"]*5

        for i in range(1):
            for j in range(5):
                label = CTkLabel(self.frameParametrePara,text=listePara[j],font=("",20))
                label.grid(row=j,column=i,pady=20)
        self.frameParametrePara1 = CTkFrame(self.frameParametre,fg_color="white")
        self.frameParametrePara1.place(relheight=1,relwidth=0.5,x=370)
        for i in range(2):
            for j in range(5):
                index = i*5+j
                
                button = CTkButton(self.frameParametrePara1,text=listebutt[index],font=("",20))
                button.grid(row=j,column=i+3,pady=20,padx=20)
    def frame2(self):
        self.tab2 = CTkFrame(self.fr2,width=750,height=750,fg_color="#FCC6FF")
        self.tab2.pack()
        
        self.frabas = CTkFrame(self.tab2,fg_color="#4B164C")
        self.frabas.place(relheight=0.1,relwidth=1,y=650)
        self.buttonRetoure = CTkButton(self.frabas,text="⬅️",width=50,height=70,font=("",50),text_color="white",fg_color="green",command=lambda:self.changerFrame(self.fr1))
        self.buttonRetoure.place(x=0,y=0)
        def k():
            self.T= True
        self.buttonRetoure = CTkButton(self.frabas,text="JOUER",width=50,height=70,font=("",50),text_color="white",fg_color="green",command=lambda:[self.changerFrame(self.fr3),k()])
        self.buttonRetoure.place(x=570,y=0)
        self.labelBas2 = CTkLabel(self.tab2,text="",width=750,fg_color="#EDF4C2")
        self.labelBas2.place(x=0,y=720)
        
        self.Niveau = CTkFrame(self.tab2,fg_color="#4B164C")
        self.Niveau.place(relheight=0.15,relwidth=1)
        
        self.labelNiveau = CTkLabel(self.Niveau,text='NIVEAU:',font=("Arial",30))
        self.labelNiveau.grid(row=0,column=0)
        
        self.listeTaille = ['4x3','4x4','5x4','6x5','6x6','7x6','8x6','9x6','10x7','10x8','12x9','12x10']
       
       
        for j in range(3):
            for i in range(4):
                index = i*3+j
                ligne,colonne = map(int, self.listeTaille[index].split('x'))
                if index < len(self.listeTaille):
                    button = CTkButton(self.Niveau, text=self.listeTaille[index],
                                        command=lambda l=ligne, c=colonne: self.changerTaillegrille(l, c))
                    button.grid(row=j,column=i+1,padx=5,pady=5)
               
        
     
                    
        frameTypeImage = CTkFrame(self.tab2, fg_color="red")
        frameTypeImage.place(relwidth=1, y=110, relheight=0.73)

        
        canvas = CTkCanvas(frameTypeImage, bg="white")
        canvas.pack(side="left", fill="both", expand=True)

        
        scrBar = CTkScrollbar(frameTypeImage, command=canvas.yview)
        scrBar.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=scrBar.set)

       
        scrollable_frame =CTkFrame(canvas)
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        def update_scrollregion(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        for i in range(5):
            for j in range(4):
                index = i*4+j
                btn = CTkButton(scrollable_frame,width=170,text="",height=150,image=self.imageschoix[index],font=("",30), compound="top"
                                ,text_color="green",fg_color="gray",hover_color="white")
                if self.listeImagechoix[index] == "voiture.jpeg":
                    btn.configure(text="VOITURE",command=lambda:self.changerImage(self.liste2))
                elif self.listeImagechoix[index] == "manga.jpeg":
                    btn.configure(text="MANGA")
                elif self.listeImagechoix[index] == "acteurs.jpeg":
                    btn.configure(text="ACTEURS")
                elif self.listeImagechoix[index] == "fruit.jpeg":
                    btn.configure(text="FRUITS")
                elif self.listeImagechoix[index] == "ville.jpeg":
                    btn.configure(text="VILLES")
                elif self.listeImagechoix[index] == "cartejeu.jpeg":
                    btn.configure(text="CARTE JEU")
                elif self.listeImagechoix[index] == "science.jpeg":
                    btn.configure(text="SCIENCE")
                elif self.listeImagechoix[index] == "sport.jpeg":
                    btn.configure(text="SPORT")
                elif self.listeImagechoix[index] == "afrique.jpeg":
                    btn.configure(text="AFRIQUE")
                elif self.listeImagechoix[index] == "musique.jpeg":
                    btn.configure(text="MUSIQUE")
                elif self.listeImagechoix[index] == "drapeau.jpeg":
                    btn.configure(text="DRAPEAUX")
                elif self.listeImagechoix[index] == "panda.jpeg":
                    btn.configure(text="ANIMAUX",command=lambda:self.changerImage(self.listeImage))
                btn.grid(row=i,column=j,padx =5,pady=5)

        scrollable_frame.bind("<Configure>", update_scrollregion)
    def changerImage(self,imagesAchanger):
        self.listeImage = imagesAchanger
        self.images.clear()
        self.im = [Image.open(img) for img in self.listeImage]
        self.images =[]
        for i in self.im:
            l = i.resize((self.imagelongeur,self.imagelargeur))
            self.images.append(l)
        self.images = [ImageTk.PhotoImage(img) for img in self.images]
            
        print("Image a changer")
        
    def frame3(self):
        self.tab3 = CTkFrame(self.fr3,width=750,height=750,fg_color="#FCC6FF")
        self.tab3.pack()
        self.frabas = CTkFrame(self.tab3,fg_color="yellow",width=750,height=45)
        self.frabas.pack()
        def Message():
            message = CTkMessagebox(title="toha",message="raouda",icon="question",option_1='OUI',option_2='NON')
            if message.get() == 'OUI':
                self.coup = 0
                self.couplabel.configure(text=f"COUPS: {self.coup}")
                self.changerFrame(self.fr2)
            else:
                pass
        self.buttonRetoure = CTkButton(self.frabas,text="⬅️",width=50,height=10,font=("",30),text_color="white",fg_color="green",command=Message)
        self.buttonRetoure.place(x=0,y=0)
        
        self.CntButton2 = CTkFrame(self.tab3,fg_color="gray",width=750,height=700)
        self.CntButton2.pack_propagate(False)
        self.CntButton2.pack()
        self.CntButton = CTkFrame(self.CntButton2,fg_color="gray",width=750,height=700)
        self.CntButton.pack(expand=True)
        self.labelBas3 = CTkLabel(self.tab3,text="",width=750,fg_color="#EDF4C2")
        self.labelBas3.place(x=0,y=720)
        self.scorelabel = CTkLabel(self.frabas,text="SCORE: 0",fg_color="white",corner_radius=5,font=("",20))
        self.scorelabel.place(x=70,y=5)
        self.erreurlabel = CTkLabel(self.frabas,text="ERREURS: 0",fg_color="white",corner_radius=5,font=("",20))
        self.erreurlabel.place(x=190,y=5)
        self.couplabel = CTkLabel(self.frabas,text="COUPS: 0",fg_color="white",corner_radius=5,font=("",20))
        self.couplabel.place(x=330,y=5)
        self.templabel = CTkLabel(self.frabas,text="00 : 00 : 00",fg_color="white",corner_radius=5,font=("",20))
        self.templabel.place(x=460,y=5)
        def pause():
            pauseFrame = CTkFrame(self.CntButton,width=500,height=500)
            pauseFrame.pack_propagate(False)
            pauseFrame.place(relx=0.5, rely=0.5, anchor="center")
            self.pause.configure(state=DISABLED)
            self.T = False
            print("pause")
            def reprise():
                pauseFrame.place(relwidth=0,relheight=0)
                self.pause.configure(state="active")
                self.T = True
            reprisreB = CTkButton(pauseFrame,command=reprise)
            reprisreB.pack(expand=True)
        self.pause = CTkButton(self.frabas,text="⏸️",font=("",30),width=50,fg_color="green",command=pause)
        self.pause.place(x=700,y=3)
        
        self.CreeGrille()
        self.temps()
    def temps(self):
        if self.T:
            self.second +=1
            self.templabel.configure(text=f"00 : 0{self.minute} : 0{self.second}")
        if self.second>=10:
            self.templabel.configure(text=f"00 : 0{self.minute} : {self.second}")
        if self.second == 60:
            self.minute += 1
            self.second = 0
            self.templabel.configure(text=f"00 : 0{self.minute} : 0{self.second}")
        self.fenetre.after(1000,lambda:self.temps())
    
    
    def CreeGrille(self):
        
        for widget in self.CntButton.winfo_children():
            widget.destroy()
        self.n = [] 
        for i in range(self.ligne):
            w = []
            for j in range(self.colonne):
                index = i*self.colonne+j
                buttonj = Button(self.CntButton,width=self.largeur,height=self.longeur,borderwidth=0,bg="blue",image=self.photo,
                                 command=lambda lig=i,col=j:self.retournementCarte(lig,col))
                buttonj.grid(row=i,column=j,padx=5,pady=5)
                w.append(buttonj)
            self.n.append(w)
    
    def changerFrame(self,frame):
        frame.tkraise()
        
    def Musique(self):
        pygame.init()
        musique = pygame.mixer.Sound("happy-kids-music-307326.mp3")
        musique.play(-1)
        
        
    def run(self):
        self.fenetre.mainloop()
