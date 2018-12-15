/// @description Insérez la description ici
// Vous pouvez écrire votre code dans cet éditeur
if (instance_exists(Highlight)) 
{
	this_highlight = instance_nearest(x,y,Highlight);
	if (moves > 0 && this.state == 1 && Case.this.stateCase == 0)
	{
		if (keyboard_check(ord("A")) || keyboard_check_pressed(ord("Z")))
		{
			if (this.x - this_highlight.x == -16 && this.y - this_highlight.y == 25)
			{
				this.x = this_highlight.x;
				this.y = this_highlight.y + 1;
				moves-- ;
				instance_destroy(instance_nearest(x,y, Chef_Highlight));
				instance_create_layer(this.x, this.y - 1, "Chef", Chef_Highlight);
			}
			else if (this.x - this_highlight.x == -32 && this.y - this_highlight.y == 1)
			{
				this.x = this_highlight.x;
				this.y = this_highlight.y + 1;
				moves-- ;
				instance_destroy(instance_nearest(x,y, Chef_Highlight));
				instance_create_layer(this.x, this.y - 1, "Chef", Chef_Highlight);
			}
			else if (this.x - this_highlight.x == -16 && this.y - this_highlight.y == -23)
			{
				this.x = this_highlight.x;
				this.y = this_highlight.y + 1;
				moves-- ;
				instance_destroy(instance_nearest(x,y, Chef_Highlight));
				instance_create_layer(this.x, this.y - 1, "Chef", Chef_Highlight);
			}
			else if (this.x - this_highlight.x == 16 && this.y - this_highlight.y == 25)
			{
				this.x = this_highlight.x;
				this.y = this_highlight.y + 1;
				moves-- ;
				instance_destroy(instance_nearest(x,y, Chef_Highlight));
				instance_create_layer(this.x, this.y - 1, "Chef", Chef_Highlight);
			}
			else if (this.x - this_highlight.x == 32 && this.y - this_highlight.y == 1)
			{
				this.x = this_highlight.x;
				this.y = this_highlight.y + 1;
				moves-- ;
				instance_destroy(instance_nearest(x,y, Chef_Highlight));
				instance_create_layer(this.x, this.y - 1, "Chef", Chef_Highlight);
			}
			else if (this.x - this_highlight.x == 16 && this.y - this_highlight.y == -23)
			{
				this.x = this_highlight.x;
				this.y = this_highlight.y + 1;
				moves-- ;
				instance_destroy(instance_nearest(x,y, Chef_Highlight));
				instance_create_layer(this.x, this.y - 1, "Chef", Chef_Highlight);
			}
		}
	}


	//var 'clickdouble': 0=noclick, 1=singleclick, 2=doubleclick
	mdoubleclick--; if(mdoubleclick<0){clickdouble=0;}
	if(mouse_check_button_pressed(mb_left) && mdoubleclick>=0 && clickdouble==0.5){clickdouble=2;}
	if(mouse_check_button_pressed(mb_left) && mdoubleclick<0){mdoubleclick=room_speed*0.25; clickdouble=0.5;}
	if(clickdouble==0.5 && mdoubleclick==0){clickdouble=1;}

	if (this.x == this_highlight.x && this.y == this_highlight.y + 1 && clickdouble == 2)
	{
		//for (i=0;i=2;i++){}
		instance_destroy(instance_nearest(x,y, Chef_Highlight));
		instance_create_layer(this.x, this.y - 1, "Chef", Chef_Highlight);
	}
	if (instance_exists(Chef_Highlight) && Chef_Highlight.x == this.x && Chef_Highlight.y == this.y - 1)
	{
		this.state = 1;
	}
	else
	{
		this.state = 0;
	}
}