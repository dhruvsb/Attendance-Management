package attpackage;

import javafx.application.Application;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.CheckBox;
import javafx.scene.control.DatePicker;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

import java.io.IOException;
import java.time.LocalDate;

import static attpackage.DateRange.DateRangeMethod;
import static attpackage.FullReport.ReportMethod;
import static attpackage.LessAttendance.LAmethod;
import static attpackage.OnlyAttendace.OnlyAttendanceMethod;

/**
 * Created by Dhruv Sb on 15-10-2016.
 */

public class Main extends Application {
    
    Stage window;
    Scene scene;
    Button GenerateButton;
    static LocalDate StartDate;
    static LocalDate EndDate;
    public static void main(String[] args) {
        launch(args);
    }
    
    @Override
    public void start(Stage primaryStage) throws Exception {
        window = primaryStage;
        window.setTitle("Attendance Management");
        
        CheckBox box1 = new CheckBox("Full attendance Sheet");
        CheckBox box2 = new CheckBox("Less attendance list");
        CheckBox box3 = new CheckBox("Only overall attendance");
        CheckBox box4 = new CheckBox("Attendance in specific period");
        box1.setSelected(true);
        DatePicker picker1 = new DatePicker();
        DatePicker picker2 = new DatePicker();
        Button SetDateButton = new Button("Set Dates");
        SetDateButton.setOnAction(event -> {
            StartDate = picker1.getValue();
            EndDate = picker2.getValue();
        });
        GenerateButton = new Button("Generate!");
        GenerateButton.setOnAction(e -> {
            try {
                handleOptions(box1, box2, box3, box4);
            } catch (IOException e1) {
                e1.printStackTrace();
            }
        });
        
        VBox layout = new VBox(15);
        layout.setPadding(new Insets(20,20, 20, 20));
        layout.getChildren().addAll(box1, box2, box3, box4,picker1,picker2,SetDateButton, GenerateButton);
        
        scene = new Scene(layout, 300, 310);
        window.setScene(scene);
        window.show();
    }

    private void handleOptions(CheckBox box1, CheckBox box2,CheckBox box3, CheckBox box4) throws IOException {
        
        if(box1.isSelected())
            ReportMethod();
        if(box2.isSelected())
            LAmethod();
        if(box3.isSelected())
            OnlyAttendanceMethod();
        if(box4.isSelected())
            DateRangeMethod(StartDate,EndDate);
    }
}