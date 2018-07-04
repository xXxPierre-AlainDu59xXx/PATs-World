/// @description Variables Chef
// Vous pouvez écrire votre code dans cet éditeur

//Variables taches
tache = 0;
efficacite = 0;
bonus = 0;

//Variables statistiques
charisme = 0;
habilite = 0;
endurance = 0;
force = 0;
niveau = charisme + habilite + endurance + force;

//Variables groupe
tailleGroupe = 0;
tailleGroupeMax = niveau*5;

//Variable sprite Chef
sprChef = 0;

//Variables déplacements
pointsDeplacements = ceil(2 + endurance/2);