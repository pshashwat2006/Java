package com.kiet.examproject;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.SQLIntegrityConstraintViolationException;
import java.util.Scanner;

/**
 * OrderManagementSystem - Java JDBC + MySQL Lab Exam Universal Solution
 * 
 * Project Package: com.kiet.examproject
 * Database: ecommerce_db
 * Table: Orders (order_id, customer_name, product_name, quantity, unit_price, status)
 * 
 * Operations Implemented:
 *  a) Database connection & insert >=3 records with duplicate key handling
 *  b) Display orders by status with quantity > 2 sorted by unit_price DESC
 *  c) Increase quantity by order ID with not-found validation
 *  d) Update status (Pending, Shipped, Delivered) only if order exists & status valid
 *  e) Display total order value (SUM) & delete record only if status is 'Cancelled'
 *  + Display all orders for live verification
 */
public class OrderManagementSystem {

    // Database configuration constants
    private static final String DB_URL = "jdbc:mysql://localhost:3306/ecommerce_db?useSSL=false&allowPublicKeyRetrieval=true";
    private static final String DB_USER = "root";
    private static final String DB_PASSWORD = "Harsh@1728";

    // Scanner for reading console inputs
    private static final Scanner scanner = new Scanner(System.in);

    public static void main(String[] args) {
        System.out.println("===============================================================");
        System.out.println("      E-COMMERCE ORDER MANAGEMENT SYSTEM (JDBC + MySQL)        ");
        System.out.println("===============================================================");

        // Step 1: Load MySQL JDBC Driver
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            System.out.println("[INFO] MySQL JDBC Driver registered successfully.");
        } catch (ClassNotFoundException e) {
            System.err.println("[ERROR] MySQL JDBC Driver not found! Make sure mysql-connector-j is in pom.xml.");
            e.printStackTrace();
            return;
        }

        // Step 2: Test Database Connection
        try (Connection conn = getConnection()) {
            System.out.println("[INFO] Connected to database 'ecommerce_db' successfully.\n");
        } catch (SQLException e) {
            System.err.println("[ERROR] Failed to connect to MySQL database! Please verify MySQL is running and credentials are correct.");
            e.printStackTrace();
            return;
        }

        // Step 3: Run Menu-Driven Loop
        boolean running = true;
        while (running) {
            printMenu();
            System.out.print("Enter your choice (1-8): ");
            String input = scanner.nextLine().trim();

            switch (input) {
                case "1":
                    // Requirement a: Insert at least three order records (Batch / Sample data)
                    insertSampleOrders();
                    break;
                case "2":
                    // Requirement a: Insert a single order record with duplicate key handling
                    insertSingleOrder();
                    break;
                case "3":
                    // Requirement b: Filter by status & quantity > 2, sort by unit_price DESC
                    displayOrdersByStatusAndQuantity();
                    break;
                case "4":
                    // Requirement c: Increase existing quantity for an order ID
                    increaseOrderQuantity();
                    break;
                case "5":
                    // Requirement d: Update status if valid (Pending, Shipped, Delivered)
                    updateOrderStatus();
                    break;
                case "6":
                    // Requirement e: Total order value (SUM) & delete if status is Cancelled
                    totalOrderValueAndDeleteCancelled();
                    break;
                case "7":
                    // Helper operation: Display all records to verify state
                    displayAllOrders();
                    break;
                case "8":
                    System.out.println("\nExiting Order Management System. Thank you!");
                    running = false;
                    break;
                default:
                    System.out.println("\n[INVALID CHOICE] Please enter a valid number between 1 and 8.\n");
            }
        }
    }

    /**
     * Helper method to establish and return a new MySQL database connection.
     */
    private static Connection getConnection() throws SQLException {
        return DriverManager.getConnection(DB_URL, DB_USER, DB_PASSWORD);
    }

    /**
     * Displays the main interactive menu to the user.
     */
    private static void printMenu() {
        System.out.println("\n----------------------------- MENU ----------------------------");
        System.out.println("1. Insert 3+ Sample Orders (PreparedStatement & Duplicate Handling) [Req a]");
        System.out.println("2. Insert a Custom Order Record (With Duplicate ID Check)           [Req a]");
        System.out.println("3. View Orders by Status (Qty > 2, Sorted by Unit Price DESC)       [Req b]");
        System.out.println("4. Increase Order Quantity by Order ID                              [Req c]");
        System.out.println("5. Update Order Status (Pending / Shipped / Delivered)             [Req d]");
        System.out.println("6. Display Total Order Value & Delete Cancelled Order               [Req e]");
        System.out.println("7. Display All Orders (Full Table View)                             [Helper]");
        System.out.println("8. Exit Application");
        System.out.println("---------------------------------------------------------------");
    }

    // =========================================================================
    // OPERATION a: Insert at least three order records using PreparedStatement.
    //             Handle duplicate order IDs appropriately.
    // =========================================================================

    /**
     * Inserts at least three predefined sample order records into the Orders table.
     * Demonstrates duplicate order_id handling using SQLIntegrityConstraintViolationException.
     */
    private static void insertSampleOrders() {
        System.out.println("\n--- [Operation a] Inserting Sample Orders ---");

        // Define at least 3 orders (5 diverse samples to test all operations)
        Object[][] sampleOrders = {
            {101, "Alice Smith", "Wireless Earbuds", 4, 1499.00, "Pending"},
            {102, "Bob Johnson", "Gaming Mouse", 5, 899.50, "Shipped"},
            {103, "Charlie Davis", "Mechanical Keyboard", 1, 3499.00, "Delivered"},
            {104, "Diana Prince", "Smart Watch", 3, 4999.00, "Pending"},
            {105, "Edward Stark", "USB-C Fast Cable", 6, 299.00, "Cancelled"}
        };

        String sql = "INSERT INTO Orders (order_id, customer_name, product_name, quantity, unit_price, status) VALUES (?, ?, ?, ?, ?, ?)";

        try (Connection conn = getConnection();
             PreparedStatement ps = conn.prepareStatement(sql)) {

            int insertedCount = 0;
            int duplicateCount = 0;

            for (Object[] order : sampleOrders) {
                int orderId = (int) order[0];
                String customerName = (String) order[1];
                String productName = (String) order[2];
                int quantity = (int) order[3];
                double unitPrice = (double) order[4];
                String status = (String) order[5];

                ps.setInt(1, orderId);
                ps.setString(2, customerName);
                ps.setString(3, productName);
                ps.setInt(4, quantity);
                ps.setDouble(5, unitPrice);
                ps.setString(6, status);

                try {
                    int rows = ps.executeUpdate();
                    if (rows > 0) {
                        System.out.println("  [SUCCESS] Inserted Order ID " + orderId + " (" + customerName + ", " + productName + ")");
                        insertedCount++;
                    }
                } catch (SQLIntegrityConstraintViolationException e) {
                    // Duplicate primary key handled gracefully
                    System.out.println("  [DUPLICATE] Order ID " + orderId + " already exists in the database. Skipped.");
                    duplicateCount++;
                }
            }

            System.out.println("\nSummary: " + insertedCount + " order(s) inserted, " + duplicateCount + " duplicate(s) handled.");

        } catch (SQLException e) {
            System.err.println("[ERROR] Database error during sample order insertion: " + e.getMessage());
        }
    }

    /**
     * Accepts a single custom order from user input and inserts using PreparedStatement.
     * Demonstrates handling duplicate primary key on user input.
     */
    private static void insertSingleOrder() {
        System.out.println("\n--- [Operation a] Insert Custom Order ---");
        try {
            System.out.print("Enter Order ID (Integer): ");
            int orderId = Integer.parseInt(scanner.nextLine().trim());

            System.out.print("Enter Customer Name: ");
            String customerName = scanner.nextLine().trim();

            System.out.print("Enter Product Name: ");
            String productName = scanner.nextLine().trim();

            System.out.print("Enter Quantity (Integer): ");
            int quantity = Integer.parseInt(scanner.nextLine().trim());

            System.out.print("Enter Unit Price (Double): ");
            double unitPrice = Double.parseDouble(scanner.nextLine().trim());

            System.out.print("Enter Status (Pending/Shipped/Delivered/Cancelled): ");
            String status = scanner.nextLine().trim();

            String sql = "INSERT INTO Orders (order_id, customer_name, product_name, quantity, unit_price, status) VALUES (?, ?, ?, ?, ?, ?)";

            try (Connection conn = getConnection();
                 PreparedStatement ps = conn.prepareStatement(sql)) {

                ps.setInt(1, orderId);
                ps.setString(2, customerName);
                ps.setString(3, productName);
                ps.setInt(4, quantity);
                ps.setDouble(5, unitPrice);
                ps.setString(6, status);

                int rows = ps.executeUpdate();
                if (rows > 0) {
                    System.out.println("\n[SUCCESS] Order ID " + orderId + " inserted successfully!");
                }
            } catch (SQLIntegrityConstraintViolationException e) {
                System.out.println("\n[DUPLICATE KEY ERROR] An order with ID " + orderId + " already exists!");
                System.out.println("Primary key violation: Duplicate order IDs are not allowed in the Orders table.");
            }

        } catch (NumberFormatException e) {
            System.out.println("\n[INPUT ERROR] Invalid number format. Please enter numeric values for ID, Quantity, and Price.");
        } catch (SQLException e) {
            System.err.println("\n[ERROR] Database error: " + e.getMessage());
        }
    }

    // =========================================================================
    // OPERATION b: Accept order status from user and display all orders with
    //             that status and quantity > 2. Sort results by unit_price DESC.
    // =========================================================================

    private static void displayOrdersByStatusAndQuantity() {
        System.out.println("\n--- [Operation b] Filter Orders by Status & Quantity > 2 ---");
        System.out.print("Enter Order Status to search (e.g., Pending, Shipped, Delivered, Cancelled): ");
        String searchStatus = scanner.nextLine().trim();

        String sql = "SELECT order_id, customer_name, product_name, quantity, unit_price, status "
                   + "FROM Orders "
                   + "WHERE LOWER(status) = LOWER(?) AND quantity > 2 "
                   + "ORDER BY unit_price DESC";

        try (Connection conn = getConnection();
             PreparedStatement ps = conn.prepareStatement(sql)) {

            ps.setString(1, searchStatus);

            try (ResultSet rs = ps.executeQuery()) {
                boolean found = false;

                while (rs.next()) {
                    if (!found) {
                        found = true;
                        System.out.println("\nOrders with Status '" + searchStatus + "' and Quantity > 2 (Sorted by Unit Price DESC):");
                        printTableHeader();
                    }
                    printTableRow(
                        rs.getInt("order_id"),
                        rs.getString("customer_name"),
                        rs.getString("product_name"),
                        rs.getInt("quantity"),
                        rs.getDouble("unit_price"),
                        rs.getString("status")
                    );
                }

                if (found) {
                    printTableFooter();
                } else {
                    System.out.println("\n[RESULT] No orders found with status '" + searchStatus + "' having quantity > 2.");
                }
            }

        } catch (SQLException e) {
            System.err.println("\n[ERROR] Database error: " + e.getMessage());
        }
    }

    // =========================================================================
    // OPERATION c: Accept order ID and additional quantity. Increase existing
    //             quantity by given amount. Display message if order not exists.
    // =========================================================================

    private static void increaseOrderQuantity() {
        System.out.println("\n--- [Operation c] Increase Order Quantity ---");
        try {
            System.out.print("Enter Order ID to update: ");
            int orderId = Integer.parseInt(scanner.nextLine().trim());

            System.out.print("Enter Additional Quantity to add: ");
            int additionalQty = Integer.parseInt(scanner.nextLine().trim());

            if (additionalQty <= 0) {
                System.out.println("[INPUT ERROR] Additional quantity must be greater than 0.");
                return;
            }

            // SQL to increase existing quantity by additional amount
            String sql = "UPDATE Orders SET quantity = quantity + ? WHERE order_id = ?";

            try (Connection conn = getConnection();
                 PreparedStatement ps = conn.prepareStatement(sql)) {

                ps.setInt(1, additionalQty);
                ps.setInt(2, orderId);

                int rowsAffected = ps.executeUpdate();

                if (rowsAffected > 0) {
                    System.out.println("\n[SUCCESS] Successfully increased quantity by " + additionalQty + " for Order ID: " + orderId);
                    // Fetch and show the updated order
                    displaySingleOrder(conn, orderId);
                } else {
                    System.out.println("\n[NOT FOUND] Order with ID " + orderId + " does not exist in the database!");
                }
            }

        } catch (NumberFormatException e) {
            System.out.println("\n[INPUT ERROR] Please enter valid integers for Order ID and Quantity.");
        } catch (SQLException e) {
            System.err.println("\n[ERROR] Database error: " + e.getMessage());
        }
    }

    // =========================================================================
    // OPERATION d: Accept order ID and new status. Retrieve order details and
    //             update status ONLY if order exists AND new status is valid
    //             (Pending, Shipped, or Delivered). Display updated details.
    // =========================================================================

    private static void updateOrderStatus() {
        System.out.println("\n--- [Operation d] Update Order Status ---");
        try {
            System.out.print("Enter Order ID to update: ");
            int orderId = Integer.parseInt(scanner.nextLine().trim());

            System.out.print("Enter New Status (Pending / Shipped / Delivered): ");
            String inputStatus = scanner.nextLine().trim();

            // Validate status: Only Pending, Shipped, Delivered are valid
            String normalizedStatus = null;
            if (inputStatus.equalsIgnoreCase("Pending")) {
                normalizedStatus = "Pending";
            } else if (inputStatus.equalsIgnoreCase("Shipped")) {
                normalizedStatus = "Shipped";
            } else if (inputStatus.equalsIgnoreCase("Delivered")) {
                normalizedStatus = "Delivered";
            } else {
                System.out.println("\n[VALIDATION ERROR] Invalid status: '" + inputStatus + "'");
                System.out.println("Allowed statuses are: Pending, Shipped, Delivered.");
                return;
            }

            try (Connection conn = getConnection()) {
                // Step 1: Check if order exists and retrieve current details
                String selectSql = "SELECT * FROM Orders WHERE order_id = ?";
                try (PreparedStatement selectPs = conn.prepareStatement(selectSql)) {
                    selectPs.setInt(1, orderId);

                    try (ResultSet rs = selectPs.executeQuery()) {
                        if (!rs.next()) {
                            System.out.println("\n[NOT FOUND] Order with ID " + orderId + " does not exist!");
                            return;
                        }

                        // Display details before update
                        System.out.println("\n--- Current Order Details (Before Update) ---");
                        System.out.println("  Order ID     : " + rs.getInt("order_id"));
                        System.out.println("  Customer     : " + rs.getString("customer_name"));
                        System.out.println("  Product      : " + rs.getString("product_name"));
                        System.out.println("  Quantity     : " + rs.getInt("quantity"));
                        System.out.println("  Unit Price   : " + rs.getDouble("unit_price"));
                        System.out.println("  Status       : " + rs.getString("status"));
                    }
                }

                // Step 2: Update the status since order exists and new status is valid
                String updateSql = "UPDATE Orders SET status = ? WHERE order_id = ?";
                try (PreparedStatement updatePs = conn.prepareStatement(updateSql)) {
                    updatePs.setString(1, normalizedStatus);
                    updatePs.setInt(2, orderId);

                    int rows = updatePs.executeUpdate();
                    if (rows > 0) {
                        System.out.println("\n[SUCCESS] Status updated to '" + normalizedStatus + "' successfully!");
                        // Step 3: Retrieve and display updated order details
                        System.out.println("\n--- Updated Order Details (After Update) ---");
                        displaySingleOrder(conn, orderId);
                    }
                }
            }

        } catch (NumberFormatException e) {
            System.out.println("\n[INPUT ERROR] Please enter a valid integer for Order ID.");
        } catch (SQLException e) {
            System.err.println("\n[ERROR] Database error: " + e.getMessage());
        }
    }

    // =========================================================================
    // OPERATION e: Display total order value using SUM(quantity * unit_price).
    //             Accept order ID and delete record ONLY if status is 'Cancelled'.
    //             Display appropriate message if order cannot be deleted.
    // =========================================================================

    private static void totalOrderValueAndDeleteCancelled() {
        System.out.println("\n--- [Operation e] Total Order Value & Conditional Delete ---");

        try (Connection conn = getConnection()) {
            // Part 1: Display total order value using SUM(quantity * unit_price)
            String sumSql = "SELECT SUM(quantity * unit_price) AS total_value FROM Orders";
            try (PreparedStatement sumPs = conn.prepareStatement(sumSql);
                 ResultSet rs = sumPs.executeQuery()) {

                if (rs.next()) {
                    double totalValue = rs.getDouble("total_value");
                    System.out.printf("[AGGREGATE RESULT] Total Order Value (SUM of quantity * unit_price) = Rs. %.2f%n", totalValue);
                }
            }

            // Part 2: Accept order ID and delete ONLY if status is 'Cancelled'
            System.out.println("\n--- Delete Order (Restricted to 'Cancelled' Status Only) ---");
            System.out.print("Enter Order ID to delete: ");
            String idInput = scanner.nextLine().trim();

            if (idInput.isEmpty()) {
                System.out.println("[INFO] No Order ID entered. Skipping delete operation.");
                return;
            }

            int orderId = Integer.parseInt(idInput);

            // Step A: Check if the order exists and retrieve its current status
            String checkSql = "SELECT status, customer_name, product_name FROM Orders WHERE order_id = ?";
            String currentStatus = null;
            String customer = null;
            String product = null;

            try (PreparedStatement checkPs = conn.prepareStatement(checkSql)) {
                checkPs.setInt(1, orderId);
                try (ResultSet rs = checkPs.executeQuery()) {
                    if (!rs.next()) {
                        System.out.println("\n[DELETE REJECTED] Order with ID " + orderId + " does not exist!");
                        return;
                    }
                    currentStatus = rs.getString("status");
                    customer = rs.getString("customer_name");
                    product = rs.getString("product_name");
                }
            }

            // Step B: Check if status is Cancelled
            if (!currentStatus.equalsIgnoreCase("Cancelled")) {
                System.out.println("\n[DELETE REJECTED] Order ID " + orderId + " cannot be deleted!");
                System.out.println("Reason: Order status is '" + currentStatus + "'. Only orders with status 'Cancelled' can be deleted.");
                return;
            }

            // Step C: Delete the cancelled order
            String deleteSql = "DELETE FROM Orders WHERE order_id = ?";
            try (PreparedStatement delPs = conn.prepareStatement(deleteSql)) {
                delPs.setInt(1, orderId);
                int rows = delPs.executeUpdate();

                if (rows > 0) {
                    System.out.println("\n[SUCCESS] Order ID " + orderId + " (" + customer + " - " + product + ") with status 'Cancelled' deleted successfully.");
                } else {
                    System.out.println("\n[ERROR] Failed to delete Order ID " + orderId + ".");
                }
            }

        } catch (NumberFormatException e) {
            System.out.println("\n[INPUT ERROR] Please enter a valid integer for Order ID.");
        } catch (SQLException e) {
            System.err.println("\n[ERROR] Database error: " + e.getMessage());
        }
    }

    // =========================================================================
    // HELPER METHODS: Display records in neat formatted tables
    // =========================================================================

    /**
     * Displays all orders present in the Orders table.
     */
    private static void displayAllOrders() {
        System.out.println("\n--- [Helper] Current Orders Table Content ---");
        String sql = "SELECT order_id, customer_name, product_name, quantity, unit_price, status FROM Orders ORDER BY order_id ASC";

        try (Connection conn = getConnection();
             PreparedStatement ps = conn.prepareStatement(sql);
             ResultSet rs = ps.executeQuery()) {

            boolean found = false;
            while (rs.next()) {
                if (!found) {
                    found = true;
                    printTableHeader();
                }
                printTableRow(
                    rs.getInt("order_id"),
                    rs.getString("customer_name"),
                    rs.getString("product_name"),
                    rs.getInt("quantity"),
                    rs.getDouble("unit_price"),
                    rs.getString("status")
                );
            }

            if (found) {
                printTableFooter();
            } else {
                System.out.println("[INFO] The Orders table is currently empty.");
            }

        } catch (SQLException e) {
            System.err.println("\n[ERROR] Database error: " + e.getMessage());
        }
    }

    /**
     * Helper to display details of a single order by ID.
     */
    private static void displaySingleOrder(Connection conn, int orderId) throws SQLException {
        String sql = "SELECT order_id, customer_name, product_name, quantity, unit_price, status FROM Orders WHERE order_id = ?";
        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            ps.setInt(1, orderId);
            try (ResultSet rs = ps.executeQuery()) {
                if (rs.next()) {
                    printTableHeader();
                    printTableRow(
                        rs.getInt("order_id"),
                        rs.getString("customer_name"),
                        rs.getString("product_name"),
                        rs.getInt("quantity"),
                        rs.getDouble("unit_price"),
                        rs.getString("status")
                    );
                    printTableFooter();
                }
            }
        }
    }

    private static void printTableHeader() {
        System.out.println("+----------+----------------------+----------------------+----------+------------+------------+");
        System.out.printf("| %-8s | %-20s | %-20s | %-8s | %-10s | %-10s |%n", "Order ID", "Customer Name", "Product Name", "Quantity", "Unit Price", "Status");
        System.out.println("+----------+----------------------+----------------------+----------+------------+------------+");
    }

    private static void printTableRow(int orderId, String customer, String product, int qty, double price, String status) {
        System.out.printf("| %-8d | %-20s | %-20s | %-8d | %10.2f | %-10s |%n", orderId, customer, product, qty, price, status);
    }

    private static void printTableFooter() {
        System.out.println("+----------+----------------------+----------------------+----------+------------+------------+");
    }
}
