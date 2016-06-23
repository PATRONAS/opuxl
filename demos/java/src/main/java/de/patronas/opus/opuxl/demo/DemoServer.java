/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.demo;

import de.patronas.opus.opuxl.FunctionHandler;
import de.patronas.opus.opuxl.OpuxlServer;

/**
 * @author stepan
 */
public class DemoServer {

  public static void main(final String[] args) {

    // Create the class which contains the methods for the demo. This methods
    // are called from within the excel sheet and return their result directly
    // to the excel sheet.
    final DemoMethods demoMethods = new DemoMethods();

    // Create a function handler and register all functions, which should be
    // available within the Excel sheet.
    final FunctionHandler functionHandler = new FunctionHandler();
    functionHandler.registerMethods(demoMethods);

    // Start the OpuxlServer with the configured function handler
    new OpuxlServer(functionHandler).start();
  }
}
