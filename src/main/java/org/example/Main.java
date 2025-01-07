package org.example;

import static org.example.Helper.*;

public class Main {
    public static void main(String[] args) {
        try {
            String exlPath = "C:\\Users\\mike\\Desktop\\data2.xlsx";
            exlToSql(exlPath);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}