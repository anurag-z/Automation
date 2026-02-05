public class TaxNavigation
{
    public void GoToFederalForms()
    {
        // STEP 1: Open Menu via Escape
        InputManager.PressKey(InputManager.SC_ESCAPE, 500);

        // STEP 2: Select 'G' for "Go To Screen"
        InputManager.PressKey(0x22, 500); // 0x22 is 'G'

        // STEP 3: Type the Screen Name
        InputManager.TypeString("FED"); 

        // STEP 4: Confirm with Enter
        InputManager.PressKey(InputManager.SC_ENTER, 1000);
    }

    public void OpenSpecificMenu(char initial)
    {
        InputManager.PressKey(InputManager.SC_ESCAPE, 300);
        InputManager.TypeString(initial.ToString());
    }
}