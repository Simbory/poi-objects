package test.drawing;

import org.apache.poi.objects.DrawingFactory;

import java.io.IOException;

public class TestModelDrawining extends DrawingFactory<TestModel> {
    public TestModelDrawining(String path) throws IllegalArgumentException, IOException {
        super(path);
    }
}