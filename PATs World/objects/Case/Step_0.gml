/// @description Méthodes Case
// Vous pouvez écrire votre code dans cet éditeur

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

if (mouse_check_button(mb_left))
{
    if((mouse_x + 13 > x) && (mouse_x - 13 < x) && (mouse_y + 13 > y) && (mouse_y - 13 < y))
    {
		instance_create_layer(x,y, "Chef", Highlight);
		stateCase = 1;
	}
}