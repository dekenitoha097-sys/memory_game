import customtkinter as ctk


ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue") 

app = ctk.CTk()
app.title("Formulaire Stylé")
app.geometry("400x400")

titre = ctk.CTkLabel(app, text="Formulaire de contact", font=("Arial", 22, "bold"))
titre.pack(pady=20)


entry_nom = ctk.CTkEntry(app, placeholder_text="Nom")
entry_nom.pack(pady=10, ipadx=10, ipady=5)

entry_email = ctk.CTkEntry(app, placeholder_text="Email")
entry_email.pack(pady=10, ipadx=10, ipady=5)


entry_mdp = ctk.CTkEntry(app, placeholder_text="Mot de passe", show="*")
entry_mdp.pack(pady=10, ipadx=10, ipady=5)


btn_envoyer = ctk.CTkButton(app, text="Envoyer")
btn_envoyer.pack(pady=20)


app.mainloop()