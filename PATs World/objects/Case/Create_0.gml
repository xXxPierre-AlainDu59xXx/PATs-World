/// @description Variables Case

//Variable d'instance
idCase = 0;
typeCase = 0;
typeStockage = 0;
nombreStockage = 0;
occupation = 0;

//Variables création map
random_set_seed(date_current_datetime());
randomize();
rand = irandom_range(1, 100);

//Variables clic case
<<<<<<< HEAD
stateCase = 0;
=======
stateCase = 0;

//Creation case
if (rand > 0 && rand <= 15)
{
	sprite_index = spr_Montagne;
}
else if (rand > 15 && rand <= 30)
{
	sprite_index = spr_Foret;
}
else if (rand > 30 && rand <= 45)
{
	sprite_index = spr_Plaine;
}
else if (rand > 45 && rand <= 60)
{
	sprite_index = spr_Riviere;
}
else if (rand > 60 && rand <= 100)
{
	sprite_index = spr_Desert;
}
>>>>>>> parent of 0b29438... oui
