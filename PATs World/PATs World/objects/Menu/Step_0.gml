/// @description Insérez la description ici
// Vous pouvez écrire votre code dans cet éditeur

instance_create_layer(x, y, "Buttons", Button_Turn);

if (keyboard_check_pressed(ord("R")))
{
	game_restart();
}