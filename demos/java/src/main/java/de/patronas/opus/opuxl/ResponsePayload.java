/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl;

/**
 * The Response Payload encapsulates a function call result which will be send
 * to the Opuxl Addin.
 * @author stepan
 */
class ResponsePayload {

  private String name;
  private String error;
  private Object[][] data;

  public ResponsePayload() {
  }

  public ResponsePayload(final String name, final String error) {
    this.name = name;
    this.error = error;
  }

  public ResponsePayload(final String name,
      final String error,
      final Object[][] data) {
    this(name, error);
    this.data = data;
  }

  /**
   * @return the name
   */
  public String getName() {
    return name;
  }

  /**
   * @param name
   *          the type to set
   */
  public void setName(final String name) {
    this.name = name;
  }

  /**
   * @return the error
   */
  public String getError() {
    return error;
  }

  /**
   * @param error
   *          the error to set
   */
  public void setError(final String error) {
    this.error = error;
  }

  /**
   * @return the data
   */
  public Object[][] getData() {
    return data;
  }

  /**
   * @param data
   *          the data to set
   */
  public void setData(final Object[][] data) {
    this.data = data;
  }

}
