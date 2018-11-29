/// @description Méthodes Case

//Highlight case
if (mouse_check_button_pressed(mb_left))
{
	if(position_meeting(mouse_x,mouse_y,Case))
	{
		instance_destroy(instance_nearest(x, y, Highlight));
		instance_create_layer(x,y, "Chef", Highlight);
		//stateCase = 1;
	}
}