
import org.junit.jupiter.api.Test;

import static org.example.Helper.*;
public class HelperTest {
    @Test
    void test(){
        // name是想要获得的图片对应的数据行中的名称
        String name = "魔门";

        saveImage(getImageFromDatabase(name));
    }
}
