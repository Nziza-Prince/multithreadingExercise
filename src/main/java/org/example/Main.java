package org.example;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class Main {
    public static void main(String[] args) {
        ExecutorService executor = Executors.newFixedThreadPool(5); // Adjust pool size as needed

        for (int i = 0; i < 10; i++) { // Simulate 10 concurrent requests
            String filePath = "employees_" + i + ".xlsx";
            executor.submit(() -> {
                EmployeeExcelExporter exporter = new EmployeeExcelExporter();
                exporter.exportEmployeesToExcel(filePath);
            });
        }
        executor.shutdown();
    }
}