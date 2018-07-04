/// @description Insérez la description ici
// Vous pouvez écrire votre code dans cet éditeur

if (mouse_check_button(mb_left))
{
    if((mouse_x + 100 > Button_Turn.x) && (mouse_x - 100 < Button_Turn.x) && (mouse_y + 50 > Button_Turn.y) && (mouse_y - 50 < Button_Turn.y))
    {
		stateButton = 1;
	}
}

if (mouse_check_button_released(mb_left))
{
	stateButton = 0;
}

if (stateButton == 0)
{
	sprite_index = spr_Button_Turn;
}

if (stateButton == 1)
{
	sprite_index = spr_Button_Turn2;
}