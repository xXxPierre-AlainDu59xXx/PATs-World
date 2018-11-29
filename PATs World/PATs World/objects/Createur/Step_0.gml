/// @description Insert description here

instance_create_layer(posx,posy,"Map",Case);

posx += 32;

if(posy <= room_height)
{
	if(posx >= room_width)
	{
		if(state = 1)
		{
			posx = 0;
		}
		else
		{
			posx = 16;
		}
			posy += 24;
			state = state * -1;
	}
}
else if (posy > room_height)
{
	instance_destroy();
}

if(number > compteur)
{
	if (posy > room_height )
	{
		instance_destroy();
	}
	else if(posy <= room_height)
	{
		instance_create_layer(posx,posy,"Map",Case);
		compteur ++;
		
		posx += 32;

		if(posx >= room_width)
		{
			if(state = 1)
			{
				posx = 32;
			}
			else
			{
				posx = 16;
			}
				posy += 24;
				state = state * -1;
		}
	}
}
/*
if(mouse = 0)
{
	instance_create_layer(x,y,"Menu",Mouse);
	mouse ++;
}*/