/// @description Méthodes Case

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

//Highlight case
if (mouse_check_button_pressed(mb_left))
{
	if(position_meeting(mouse_x,mouse_y,Case))
	{
		instance_destroy(instance_nearest(x, y, Highlight));
		instance_create_layer(x,y, "Chef", Highlight);
		stateCase = 1;
	}
}