package imagesplit;

import java.util.Enumeration;
import javax.swing.AbstractButton;
import javax.swing.ButtonGroup;

/**
 *
 * @author Artur
 */
public class GroupButtonUtils {

    /**
     *
     * @param buttonGroup
     * @return
     */
    public String getSelectedButtonName(ButtonGroup buttonGroup) {
        for (Enumeration<AbstractButton> buttons = buttonGroup.getElements(); buttons.hasMoreElements();) {
            AbstractButton button = buttons.nextElement();

            if (button.isSelected()) {
                return button.getName();
            }
        }

        return null;
    }
}
