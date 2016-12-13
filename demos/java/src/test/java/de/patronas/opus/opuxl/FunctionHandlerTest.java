/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl;

import static org.junit.Assert.assertEquals;

import java.util.List;

import org.junit.Test;

import com.google.common.collect.Lists;

import de.patronas.opus.opuxl.server.ExcelArgumentAttribute;
import de.patronas.opus.opuxl.server.ExcelType;
import de.patronas.opus.opuxl.server.FunctionHandler;
import de.patronas.opus.opuxl.server.FunctionPayload;
import de.patronas.opus.opuxl.server.InstanceMethod;
import de.patronas.opus.opuxl.server.OpuxlMatrix;

/**
 * @author stepan
 */
public class FunctionHandlerTest {

  @Test
  public void testRegisterAndExecuteFunctionWithNoArgs() {
    // Given: a function handler with the `getFoobar` method registered
    final FunctionHandler functionHandler = new FunctionHandler();
    functionHandler.registerMethods(new TestMethods());

    // When: we execute the function with the function handler
    final OpuxlMatrix result = functionHandler
        .execute(new FunctionPayload("getFoobar"));

    // Then: we expect that he method is called and the resulting matrix
    // contains its result value.
    assertEquals("Expected the resulting matrix to contain foobar as it first element", "Foobar", result
        .getData().get(0).get(0));
  }

  @Test
  public void testRegisterWithAnnotation() {
    final FunctionHandler functionHandler = new FunctionHandler();
    functionHandler.registerMethods(new TestMethods());

    assertEquals("Expected 2 methods to be registered as two methods in our demoMethods class have the Opuxl Annotation", 2, functionHandler
        .getMethods().size());
  }

  @Test
  public void testRegisterWithArguments() {
    final FunctionHandler functionHandler = new FunctionHandler();
    functionHandler.registerMethods(new TestMethods());

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
    functionHandler.registerMethods(new TestMethods());

    // When: we execute the function with the function handler and a valid
    // agument
    final OpuxlMatrix result = functionHandler
        .execute(new FunctionPayload("square", Lists.<Object> newArrayList(5)));

    // Then: we expect that the input was squared and the result is present as
    // the first value in the matrix.
    assertEquals("Expected that the first entry of the matrix contains 5*5=25", 25.0, result
        .getData().get(0).get(0));
  }
}
