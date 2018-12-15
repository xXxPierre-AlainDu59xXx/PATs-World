/// @description Méthodes Case

//Highlight case
if (mouse_check_button_pressed(mb_left))
{
	if(position_meeting(mouse_x,mouse_y,this))
	{
		instance_destroy(instance_nearest(x,y, Highlight));
		instance_create_layer(x,y, "Chef", Highlight);
	}
}

if (Chef.x == this.x && Chef.y == this.y + 1)
{
	this.stateCase = 1;
}
else
{
	this.stateCase = 0;
}