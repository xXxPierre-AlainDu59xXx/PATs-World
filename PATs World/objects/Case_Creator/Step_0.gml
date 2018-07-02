/// @description Insérez la description ici
// Vous pouvez écrire votre code dans cet éditeur

if (instance_nearest(x,y, Case).Fait == 0)
{
	rand = random_range(0, 4);
	if (rand == 0)
	{
		object_set_sprite(Case, spr_Montagne);
	}
	else if (rand == 1)
	{
		object_set_sprite(Case, spr_Foret);
	}
	else if (rand == 2)
	{
		object_set_sprite(Case, spr_Plaine);
	}
	else if (rand == 3)
	{
		object_set_sprite(Case, spr_Riviere);
	}
	else if (rand == 4)
	{
		object_set_sprite(Case, spr_Desert);
	}
}