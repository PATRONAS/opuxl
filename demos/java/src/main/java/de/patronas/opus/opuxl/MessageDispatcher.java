/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.net.Socket;
import java.util.Map;
import java.util.Map.Entry;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;

/**
 * The Message Dispatcher is responsible to handle received messages and send
 * suitable results.
 * @author stepan
 */
class MessageDispatcher extends Thread {

  private static final Logger LOG = LoggerFactory
      .getLogger(MessageDispatcher.class);

  /**
   * The Initial Request from the Opuxl Addin will contain this function name
   * and expects a response with all available methods from the function
   * handler.
   */
  private static final String INITIALIZE = "INITIALIZE";

  /**
   * The Socket of the current connection. Will be closed after we send the
   * response back to the Addin.
   */
  private final Socket socket;

  private final FunctionHandler functionHandler;
  private final ObjectMapper mapper = new ObjectMapper();

  public MessageDispatcher(final Socket socket,
      final FunctionHandler functionHandler) {
    this.socket = socket;
    this.functionHandler = functionHandler;
  }

  /**
   * Handles an accepted socket connection.
   * <p>
   * Read Message -> Dispatch -> Send result
   */
  @Override
  public void start() {
    try (final BufferedReader reader = new BufferedReader(new InputStreamReader(socket
        .getInputStream()));
        final PrintWriter writer = new PrintWriter(new OutputStreamWriter(socket
            .getOutputStream()))) {
      // Read the incoming message
      final String message = readMessage(reader);
      // ...dispatch it
      final String response = dispatch(message);
      // ...and send the result back.
      sendMessage(response, writer);
    } catch (final IOException e) {
      LOG.error("Error while handling message.");
      throw new RuntimeException("Error while handling message", e);
    } finally {
      if (socket != null) {
        try {
          socket.close();
        } catch (final IOException e) {
          LOG.error("Error while closing the socket: {}", e);
        }
      }
    }
  }

  /**
   * Dispatches the given message and returns the appropriate result as a json
   * string.
   * @param message
   *          the message
   * @return the json result
   */
  private String dispatch(final String message) {
    String result = "{}";
    try {
      final FunctionPayload functionPayload = mapper
          .readValue(message, FunctionPayload.class);
      if (INITIALIZE.equals(functionPayload.getName())) {
        // On the initial request, we return a list of all available methods,
        // which will be injected into the excel sheet.
        result = createFunctionDescriptionResponse();
      } else if (functionHandler.hasFunction(functionPayload.getName())) {
        // If our function handler has a method registered for the requested
        // action, we try to execute it.
        result = executeFunction(functionPayload);
      } else {
        // If no method can be found, create an error response.
        result = createResult(functionPayload.getName(), new Object[][] {}, "Can't dispatch message. Function is not registered.");
      }
    } catch (final IOException e) {
      LOG.error("Error while dispatching message.");
      throw new RuntimeException("Error while dispatching message", e);
    }
    return result;
  }

  /**
   * Create a json result with the given arguments.
   * @param type
   *          the type of response
   * @param data
   *          the result data
   * @param error
   *          the error, if present
   * @return the result as a json string
   */
  private String createResult(final String type, final Object[][] data, final String error) {
    String result = "{}";
    final ResponsePayload payload = new ResponsePayload(type, error, data);

    try {
      result = mapper.writeValueAsString(payload);
    } catch (final JsonProcessingException ex) {
      LOG.warn("Could not create result json: {}", ex);
    }

    return result;
  }

  /**
   * Executes the function which is described by the function payload and
   * returns the json result string.
   * @param functionPayload
   *          the function payload
   * @return result as json string
   */
  private String executeFunction(final FunctionPayload functionPayload) {
    Object[][] executed = new Object[][] {};
    String error = "";
    try {
      executed = functionHandler.execute(functionPayload);
    } catch (final Exception ex) {
      error = ex.getMessage();
    }
    return createResult(functionPayload.getName(), executed, error);
  }

  /**
   * Create the Function Description Response which will register the Functions
   * of the FunctionHandler in the Excel Sheet.
   * @return the function description as a json string
   */
  private String createFunctionDescriptionResponse() {
    final Map<String, InstanceMethod> methodsMap = functionHandler.getMethods();

    final Object[][] data = new Object[methodsMap.keySet().size()][2];

    int i = 0;
    for (final Entry<String, InstanceMethod> entry : methodsMap.entrySet()) {
      // map List of ExcelArgumentAttributes to Array of String Arrays
      // containing Name+Description.

      data[i] = new Object[] {
          entry.getKey(), entry.getValue().toJsonArrayString().toArray()
      };
      i++;
    }

    return createResult(INITIALIZE, data, "");
  }

  private void sendMessage(final String message, final PrintWriter writer) {
    LOG.info("Sending message to client: [{}] {}", socket.getInetAddress(), message);
    writer.println(message);
    writer.flush();
  }

  private String readMessage(final BufferedReader reader) throws IOException {
    String message = "";
    message = reader.readLine();
    LOG.info("Received message from client: [{}]  {}", socket.getInetAddress()
        .toString(), message);
    return message;
  }

}
