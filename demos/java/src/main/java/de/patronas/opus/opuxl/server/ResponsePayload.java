/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.server;

/**
 * The Response Payload encapsulates a function call result which will be send
 * to the Opuxl Addin.
 * @author stepan
 */
public class ResponsePayload {

  private String name;
  private String error;
  private OpuxlMatrix matrix;

  /**
   * Constructor.
   */
  public ResponsePayload() {
    // For Serialization
  }

  /**
   * Constructor.
   * @param name
   *          the name
   * @param error
   *          the error
   */
  public ResponsePayload(final String name, final String error) {
    this.name = name;
    this.error = error;
  }

  /**
   * Constructor.
   * @param name
   *          the name
   * @param error
   *          the error
   * @param matrix
   *          the result data
   */
  public ResponsePayload(final String name,
      final String error,
      final OpuxlMatrix matrix) {
    this(name, error);
    this.matrix = matrix;
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
  public OpuxlMatrix getMatrix() {
    return matrix;
  }

  /**
   * @param matrix
   *          the data to set
   */
  public void setMatrix(final OpuxlMatrix matrix) {
    this.matrix = matrix;
  }

}
