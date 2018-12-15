/// @description Insert description here
//Highlight case
/*
if(position_meeting(mouse_x,mouse_y,Case) && mouse_check_button_pressed(mb_left))
{
	(instance_nearest(mouse_x,mouse_y,Case)).stateCase ++;
}

if((instance_nearest(mouse_x,mouse_y,Case)).stateCase = 1)
{
	instance_create_layer((instance_nearest(mouse_x,mouse_y,Case)).x,(instance_nearest(mouse_x,mouse_y,Case)).y, "Chef", Highlight);
}
else if((instance_nearest(mouse_x,mouse_y,Case)).stateCase = 2)
{
	instance_destroy(instance_nearest(mouse_x,mouse_y,Highlight));
	//(instance_nearest(mouse_x,mouse_y,Case)).stateCase -= 2;
}
**/