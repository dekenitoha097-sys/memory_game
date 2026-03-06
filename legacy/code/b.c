#include <stdio.h>
#include <string.h>
/*
typedef struct Employe {
    char nom[30];
    char prenom[30];
    char date[30];
    char address[30];
    char numero[30];
    int salaire;
}Employe;

float salaire(Employe Employer[],int nbr){
    float k = 0;
    for(int i=0;i<nbr;i++){
        k+=Employer[i].salaire;
    }
    return k/nbr;
}

int calcul(int a, int b){
    return a*b;
}
int main(void) {
    int nbr;

    printf("saisir le nombre s'employe:  ");
    scanf("%d",&nbr);
    Employe t[nbr];
    for(int i=0;i<nbr;i++){
        Employe p1;
        printf("saisir les info: %d\n", i+1);
        printf("nom: %d\n", i+1);
        scanf("%s",&p1.nom);
        printf("prenom: %d\n", i+1);
        scanf("%s",&p1.prenom);
        printf("date de naissace: %d\n", i+1);
        scanf("%s",&p1.date);
        printf("address : %d\n", i+1);
        scanf("%s",&p1.address);
        printf("numero de telephone %d\n", i+1);
        scanf("%s",&p1.numero);
        printf("Salaire %d\n", i+1);
        scanf("%d",&p1.salaire);
        t[i] = p1;
    }

    int a = calcul(3,3);

    printf("%d\n",a);
    printf("le salaire moyen est : %.2f",salaire(t,nbr));
    return 0;
}
*/

// %zu
/*
struct stylo{
    char marque[20];
    int couleur;
    float prix_stylo;
};
typedef struct stylo stylo;

struct livre{
    char titre[20];
    char auteur[30];
    int nb_page;
    float prix_livre;
};
typedef struct livre livre;
struct article{
    stylo stylo;
    livre livre;
   
};
typedef struct article article;


int main(){

    article a;
    strcpy(a.stylo.marque,"bic");

    a.stylo.couleur  = 1 ;

    a.stylo.prix_stylo = 1.99;


    printf("%s\n",a.stylo.marque);

    printf("%d\n",a.stylo.couleur);

    printf("%f\n",a.stylo.prix_stylo );

    article a1;
    strcpy(a1.livre.titre,"le CID");
    strcpy(a1.livre.auteur,"TOHA");
    a1.livre.nb_page = 200;
    a1.livre.prix_livre = 19.99;

    printf("%s\n",a1.livre.titre);

    printf("%s\n",a1.livre.auteur);

    printf("%d\n",a1.livre.nb_page);

    printf("%f\n",a1.livre.prix_livre);

    

    
    return 0 ;
}

*/
int f(){
    printf("hello world");
}
int main(){
    char tableau[5] = "toha";
    tableau[1] = 'w';
    printf("%s",tableau);

    f();

    return 0;
}