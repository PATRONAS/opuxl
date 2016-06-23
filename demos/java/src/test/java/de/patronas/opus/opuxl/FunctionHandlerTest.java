/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl;

import static org.junit.Assert.assertEquals;

import java.util.List;

import org.junit.Test;

import com.google.common.collect.Lists;

/**
 * @author stepan
 */
public class FunctionHandlerTest {

  @Test
  public void testRegisterAndExecuteFunctionWithNoArgs() {
    // Given: a function handler with the `getFoobar` method registered
    final FunctionHandler functionHandler = new FunctionHandler();
    functionHandler.registerMethods(new DefaultMethods());

    // When: we execute the function with the function handler
    final Object[][] result = functionHandler
        .execute(new FunctionPayload("getFoobar"));

    // Then: we expect that he method is called and the resulting matrix
    // contains its result value.
    assertEquals("Expected the resulting matrix to contain foobar as it first element", "Foobar", result[0][0]);
  }

  @Test
  public void testRegisterWithAnnotation() {
    final FunctionHandler functionHandler = new FunctionHandler();
    functionHandler.registerMethods(new DefaultMethods());

    assertEquals("Expected 2 methods to be registered as two methods in our demoMethods class have the Opuxl Annotation", 2, functionHandler
        .getMethods().size());
  }

  @Test
  public void testRegisterWithArguments() {
    final FunctionHandler functionHandler = new FunctionHandler();
    functionHandler.registerMethods(new DefaultMethods());

    final InstanceMethod squareMethod = functionHandler.getMethods()
        .get("square");

    final List<ExcelArgumentAttribute> parameterArguments = squareMethod
        .getParameterArguments();
    assertEquals(1, parameterArguments.size());
    assertEquals("Number", parameterArguments.get(0).getName());
    assertEquals("The Number which will be squared.", parameterArguments.get(0)
        .getDescription());
    assertEquals(ExcelType.NUMBER, parameterArguments.get(0).getType());
  }

  @Test
  public void testRegisterAndExecuteFunctionWithArgs() {
    // Given: a function handler with the `square` method registered
    final FunctionHandler functionHandler = new FunctionHandler();
    functionHandler.registerMethods(new DefaultMethods());

    // When: we execute the function with the function handler and a valid
    // agument
    final Object[][] result = functionHandler
        .execute(new FunctionPayload("square", Lists.<Object> newArrayList(5)));

    // Then: we expect that the input was squared and the result is present as
    // the first value in the matrix.
    assertEquals("Expected that the first entry of the matrix contains 5*5=25", 25.0, result[0][0]);
  }
}

class DefaultMethods {

  @OpuxlMethod
  public String getFoobar() {
    return "Foobar";
  }

  @OpuxlMethod
  public double square(@OpuxlArg(name = "Number", description = "The Number which will be squared.", type = ExcelType.NUMBER) final double number) {
    return Math.pow(number, 2);
  }
}