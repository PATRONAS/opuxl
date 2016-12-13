/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.server;

import java.util.ArrayList;
import java.util.List;

/**
 * The function payload will be send from the Opuxl Addin and describe how a
 * specific function should be called.
 * @author stepan
 */
public class FunctionPayload {

  private String name;
  private List<Object> args = new ArrayList<>();

  public FunctionPayload() {
  }

  public FunctionPayload(final String name) {
    this.name = name;
  }

  public FunctionPayload(final String name, final List<Object> args) {
    this(name);
    this.args = args;
  }

  /**
   * @return the name
   */
  public String getName() {
    return name;
  }

  /**
   * @param name
   *          the name to set
   */
  public void setName(final String name) {
    this.name = name;
  }

  /**
   * @return the args
   */
  public List<Object> getArgs() {
    return args;
  }

  /**
   * @param args
   *          the args to set
   */
  public void setArgs(final List<Object> args) {
    this.args = args;
  }
}
