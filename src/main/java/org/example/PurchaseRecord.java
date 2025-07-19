package org.example;

import java.time.LocalDate;
import java.math.BigDecimal;

/**
 * DTO representing a single purchase record from Excel.
 * This class holds all the data for one row and provides clean access to it.
 */
public class PurchaseRecord {
    private String itemName;
    private BigDecimal price;
    private int quantity;
    private LocalDate purchaseDate;
    private String category;
    private String vendor;
    private BigDecimal totalCost;

    // Default constructor
    public PurchaseRecord() {
    }

    // Constructor with all fields
    public PurchaseRecord(String itemName, BigDecimal price, int quantity,
                          LocalDate purchaseDate, String category, String vendor, BigDecimal totalCost) {
        this.itemName = itemName;
        this.price = price;
        this.quantity = quantity;
        this.purchaseDate = purchaseDate;
        this.category = category;
        this.vendor = vendor;
        this.totalCost = totalCost;
    }

    // Getters and Setters
    public String getItemName() {
        return itemName;
    }

    public void setItemName(String itemName) {
        this.itemName = itemName;
    }

    public BigDecimal getPrice() {
        return price;
    }

    public void setPrice(BigDecimal price) {
        this.price = price;
    }

    public int getQuantity() {
        return quantity;
    }

    public void setQuantity(int quantity) {
        this.quantity = quantity;
    }

    public LocalDate getPurchaseDate() {
        return purchaseDate;
    }

    public void setPurchaseDate(LocalDate purchaseDate) {
        this.purchaseDate = purchaseDate;
    }

    public String getCategory() {
        return category;
    }

    public void setCategory(String category) {
        this.category = category;
    }

    public String getVendor() {
        return vendor;
    }

    public void setVendor(String vendor) {
        this.vendor = vendor;
    }

    public BigDecimal getTotalCost() {
        return totalCost;
    }

    public void setTotalCost(BigDecimal totalCost) {
        this.totalCost = totalCost;
    }

    /**
     * Method to convert to Object array for JTable.
     * Order matches the GUI table headers: Product Name, Unit Price, Qty Sold, Sale Date, Category, Customer Name, Total Amount
     */
    public Object[] toObjectArray() {
        return new Object[]{
                itemName,           // Product Name
                price,              // Unit Price
                quantity,           // Qty Sold
                purchaseDate,       // Sale Date
                category,           // Category
                vendor,             // Customer Name
                totalCost           // Total Amount
        };
    }

    // Override toString for debugging
    @Override
    public String toString() {
        return "PurchaseRecord{" +
                "itemName='" + itemName + '\'' +
                ", price=" + price +
                ", quantity=" + quantity +
                ", purchaseDate=" + purchaseDate +
                ", category='" + category + '\'' +
                ", vendor='" + vendor + '\'' +
                ", totalCost=" + totalCost +
                '}';
    }
}