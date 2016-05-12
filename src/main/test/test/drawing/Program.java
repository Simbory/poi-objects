package test.drawing;

import org.apache.poi.objects.DrawingFactory;

import java.io.IOException;
import java.util.ArrayList;

public class Program {

    public static void main(String[] args) {
        ArrayList<TestModel> models = new ArrayList<TestModel>();
        for (int i = 0; i < 500; i++) {
            TestModel m = new TestModel();
            m.setNum(i);
            m.setName("Model " + i);
            models.add(m);
        }
        try {
            DrawingFactory<TestModel> drawingFactory = new TestModelDrawining("C:\\test.xlsx");
            drawingFactory.draw("models", models);
            drawingFactory.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}