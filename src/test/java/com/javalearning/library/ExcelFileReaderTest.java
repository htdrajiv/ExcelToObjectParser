package com.javalearning.library;

import com.javalearning.library.model.Address;
import com.javalearning.library.model.Customer;
import org.junit.Assert;
import org.junit.Test;

import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelFileReaderTest {
    @Test
    public void TestFileReaderParse() {
        Map<String, String> sheetNameToClassNameMap = new HashMap<>();
        sheetNameToClassNameMap.put("customer","customer");
        sheetNameToClassNameMap.put("address","address");
        sheetNameToClassNameMap.put("phone","phone");
        ExcelFileReader fileReader = new ExcelFileReader(sheetNameToClassNameMap);

        ClassLoader classLoader = getClass().getClassLoader();
        String excelFilePath = new File(classLoader.getResource("customer.xlsx").getFile()).getPath();

        List<Customer> customerList = fileReader.parse(excelFilePath, "customer", Customer.class);
        List<Address> addressList = fileReader.parse(excelFilePath, "address", Address.class);


        Assert.assertTrue(customerList.size() > 1);

        Customer customer = customerList.get(0);
        Assert.assertEquals("customer1", customer.getName());
        Assert.assertEquals(30, customer.getAge());


        Assert.assertEquals(7, addressList.size());
        Assert.assertTrue(addressList.size() > 1);

        Address address = addressList.get(0);
        Assert.assertEquals("12345", address.getZipCode());
        Assert.assertEquals("US", address.getCountry());

    }
}
