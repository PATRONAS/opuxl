/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.server;

import com.fasterxml.jackson.annotation.JsonProperty;

/**
 * @author stepan
 */
public class ExcelArgumentAttribute {

  private String name;
  private String description;
  private ExcelType type;
  private boolean optional;

  /**
   * Constructor.
   */
  public ExcelArgumentAttribute() {
    // For Serialization
  }

  /**
   * Constructor.
   * @param name
   *          the name
   * @param description
   *          the description
   * @param type
   *          the type
   * @param optional
   *          whether the arg is optional
   */
  public ExcelArgumentAttribute(final String name,
      final String description,
      final ExcelType type,
      final boolean optional) {
    this.name = name;
    this.description = description;
    this.type = type;
    setOptional(optional);
  }

  /**
   * Constructor.
   * @param arg
   *          the opuxlArg
   */
  public ExcelArgumentAttribute(final OpuxlArg arg) {
    this(arg.name(), arg.description(), arg.type(), arg.optional());
  }

  /**
   * @return the name
   */
  @JsonProperty("Name")
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
   * @return the description
   */
  @JsonProperty("Description")
  public String getDescription() {
    return description;
  }

  /**
   * @param description
   *          the description to set
   */
  public void setDescription(final String description) {
    this.description = description;
  }

  /**
   * @return the type
   */
  @JsonProperty("Type")
  public ExcelType getType() {
    return type;
  }

  /**
   * @param type
   *          the type to set
   */
  public void setType(final ExcelType type) {
    this.type = type;
  }

  /**
   * @return the optional
   */
  @JsonProperty("Optional")
  public boolean isOptional() {
    return optional;
  }

  /**
   * @param optional
   *          the optional to set
   */
  public void setOptional(final boolean optional) {
    this.optional = optional;
  }

}
