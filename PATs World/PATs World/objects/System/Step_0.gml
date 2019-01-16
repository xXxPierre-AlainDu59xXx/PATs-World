/// @description Insérez la description ici
// Vous pouvez écrire votre code dans cet éditeur
if (keyboard_check_pressed(ord("X")) && window_get_fullscreen())
{
	window_set_fullscreen(false);
}
else if (keyboard_check_pressed(ord("X")) && window_get_fullscreen() == 0)
{
	window_set_fullscreen(true);
}
if keyboard_check_pressed(ord("R")) 
{
	game_restart();
}