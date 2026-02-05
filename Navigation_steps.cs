public class TaxFlow
{
    public void NavigateToClientSearch()
    {
        // 1. Enter the Federal Module
        InputManager.PressFKey(3); 

        // 2. Open the 'G'o-to menu via Escape + Initial
        InputManager.PressSpecialKey("ESC");
        InputManager.TypeString("G");

        // 3. Type the screen code
        InputManager.TypeString("1040");

        // 4. Submit
        InputManager.PressSpecialKey("ENTER");
    }
}