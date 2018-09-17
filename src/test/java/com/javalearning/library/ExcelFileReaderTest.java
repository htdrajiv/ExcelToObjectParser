package com.javalearning.library;

import com.javalearning.library.model.Address;
import com.javalearning.library.model.Customer;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.util.List;

public class ExcelFileReaderTest {
    @Test
    public void TestFileReaderParse() {
        ExcelFileReaderMoreAdvance fileReader = new ExcelFileReaderMoreAdvance();

        ClassLoader classLoader = getClass().getClassLoader();
        String excelFilePath = new File(classLoader.getResource("customer.xlsx").getFile()).getPath();

        List<Customer> customerList = fileReader.parse(excelFilePath, "customer", Customer.class);
        List<Address> addressList = fileReader.parse(excelFilePath, "address", Address.class);


        Assert.assertEquals(1,customerList.size());
        Assert.assertFalse(customerList.size() > 1);

        Customer customer = customerList.get(0);
        Assert.assertEquals("customer1", customer.getName());
        Assert.assertEquals(30, customer.getAge());


        Assert.assertEquals(1, addressList.size());
        Assert.assertFalse(addressList.size() > 1);

        Address address = addressList.get(0);
        Assert.assertEquals("50266", address.getZipCode());
        Assert.assertEquals("US", address.getCountry());

    }
}
