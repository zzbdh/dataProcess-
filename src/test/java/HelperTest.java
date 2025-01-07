
import org.junit.jupiter.api.Test;

import static org.example.Helper.*;
import static org.junit.jupiter.api.Assertions.*;
public class HelperTest {
    @Test
    void test(){
        // 输入是想要获得的图片的名称
        saveImage(getImageFromDatabase("suaa"));
    }
}
