/// @description Insérez la description ici
// Vous pouvez écrire votre code dans cet éditeur
if (keyboard_check_pressed(ord("P")))
{
	Chef.moves = Chef.maxmoves;
}
if keyboard_check_pressed(ord("R")) 
{
	game_restart();
}