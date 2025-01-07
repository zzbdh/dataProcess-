package org.example;

import static org.example.Helper.*;

public class Main {
    public static void main(String[] args) {
        try {
            exlToSql("C:\\Users\\mike\\Desktop\\data2.xlsx");
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}