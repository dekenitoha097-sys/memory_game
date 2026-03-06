
CREATE DATABASE MemoryGame;
USE MemoryGame;


CREATE TABLE Joueur (
    id_joueur INT AUTO_INCREMENT PRIMARY KEY,
    nom VARCHAR(50),
    type ENUM('humain', 'ordinateur') NOT NULL
);

CREATE TABLE Partie (
    id_partie INT AUTO_INCREMENT PRIMARY KEY,
    mode ENUM('solo', 'multijoueur', 'vs_ordi') NOT NULL,
    date_heure DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE Carte (
    id_carte INT AUTO_INCREMENT PRIMARY KEY,
    symbole CHAR(1) NOT NULL,
    image_path VARCHAR(100)
);

CREATE TABLE Score (
    id_score INT AUTO_INCREMENT PRIMARY KEY,
    id_partie INT,
    id_joueur INT,
    points INT DEFAULT 0,
    FOREIGN KEY (id_partie) REFERENCES Partie(id_partie),
    FOREIGN KEY (id_joueur) REFERENCES Joueur(id_joueur)
);

CREATE TABLE Coup (
    id_coup INT AUTO_INCREMENT PRIMARY KEY,
    id_partie INT,
    id_joueur INT,
    id_carte INT,
    position_lig INT,
    position_col INT,
    date_heure DATETIME DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (id_partie) REFERENCES Partie(id_partie),
    FOREIGN KEY (id_joueur) REFERENCES Joueur(id_joueur),
    FOREIGN KEY (id_carte) REFERENCES Carte(id_carte)
);


INSERT INTO Joueur (nom, type) VALUES ('John', 'humain');
INSERT INTO Joueur (nom, type) VALUES ('Ordi', 'ordinateur');



INSERT INTO Partie (mode) VALUES ('multijoueur');


INSERT INTO Carte (symbole, image_path) VALUES 
('A', 'images/A.png'),
('B', 'images/B.png'),
('C', 'images/C.png'),
('D', 'images/D.png'),
('E', 'images/E.png'),
('F', 'images/F.png'),
('G', 'images/G.png'),
('H', 'images/H.png');


INSERT INTO Score (id_partie, id_joueur, points) VALUES (1, 1, 3);
INSERT INTO Score (id_partie, id_joueur, points) VALUES (1, 2, 2);

-- 7. Insertion de coups (exemple)
INSERT INTO Coup (id_partie, id_joueur, id_carte, position_lig, position_col) VALUES (1, 1, 1, 0, 0);
INSERT INTO Coup (id_partie, id_joueur, id_carte, position_lig, position_col) VALUES (1, 2, 2, 0, 1);





CREATE TABLE joueur(
   id_joueur INT,
   nom VARCHAR(50),
   prenom VARCHAR(50),
   E_mail VARCHAR(50),
   type VARCHAR(50),
   PRIMARY KEY(id_joueur)
);

CREATE TABLE Parties(
   id_partie INT,
   date_heure DATETIME,
   Mode VARCHAR(50),
   PRIMARY KEY(id_partie)
);

CREATE TABLE Carte(
   id_carte INT,
   nom_carte VARCHAR(50),
   image_chemin INT,
   PRIMARY KEY(id_carte)
);

CREATE TABLE score(
   id_score VARCHAR(50),
   id_partie INT,
   id_joueur INT,
   points INT,
   id_partie_1 INT NOT NULL,
   id_joueur_1 INT NOT NULL,
   PRIMARY KEY(id_score),
   FOREIGN KEY(id_partie_1) REFERENCES Parties(id_partie),
   FOREIGN KEY(id_joueur_1) REFERENCES joueur(id_joueur)
);

CREATE TABLE Coup(
   id_coup INT,
   id_partie INT,
   id_joueur_ INT,
   _id_carte INT,
   position_ligne INT,
   positioin_colonne INT,
   date_heure DATETIME,
   id_partie_1 INT NOT NULL,
   id_carte INT NOT NULL,
   PRIMARY KEY(id_coup),
   FOREIGN KEY(id_partie_1) REFERENCES Parties(id_partie),
   FOREIGN KEY(id_carte) REFERENCES Carte(id_carte)
);

CREATE TABLE effectuer(
   id_joueur INT,
   id_coup INT,
   PRIMARY KEY(id_joueur, id_coup),
   FOREIGN KEY(id_joueur) REFERENCES joueur(id_joueur),
   FOREIGN KEY(id_coup) REFERENCES Coup(id_coup)
);
