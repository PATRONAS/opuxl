/*******************************************************************************
 * Copyright 2016 PATRONAS Financial Systems GmbH. All rights reserved.
 ******************************************************************************/
package de.patronas.opus.opuxl.server;

import java.lang.annotation.Annotation;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.google.common.collect.Lists;

/**
 * The function handler is responsible to store the registered functions and to
 * execute them.
 * @author stepan
 */
public class FunctionHandler {

  private static final Logger LOG = LoggerFactory
      .getLogger(FunctionHandler.class);

  private final Map<String, InstanceMethod> methods = new HashMap<>();

  /**
   * Registered all methods of the given object instances which are flagged with
   * the opuxl annotation.
   * @param instance
   *          the object instance which contains the methods
   */
  public void registerMethods(final Object instance) {
    final Method[] instanceMethods = instance.getClass().getMethods();
    // Iterate over the methods of the given object instance and register all
    // 'Opuxl' marked methods.
    for (final Method method : instanceMethods) {
      final OpuxlMethod annotation = method.getAnnotation(OpuxlMethod.class);
      if (annotation != null) {
        registerMethod(instance, method, annotation);
      }
    }
  }

  private void registerMethod(final Object instance, final Method method, final OpuxlMethod annotation) {
    final List<ExcelArgumentAttribute> arguments = Lists.newArrayList();

    final Annotation[][] parameterAnnotations = method
        .getParameterAnnotations();
    for (final Annotation[] annotations : parameterAnnotations) {
      arguments.addAll(registerParameter(annotations));
    }

    final String registeredName = annotation.name().isEmpty() ? method
        .getName() : annotation.name();

    if (hasFunction(registeredName)) {
      LOG.error("Method with name {} is already registered, not doing an override.", registeredName);
    } else {
      getMethods()
          .put(registeredName, new InstanceMethod(instance, method, annotation.description(), arguments));
    }
  }

  private List<ExcelArgumentAttribute> registerParameter(final Annotation[] annotations) {
    final List<ExcelArgumentAttribute> result = Lists.newArrayList();
    for (final Annotation paramAnnotation : annotations) {
      // Extract the information from the Opuxl Argument Annotation which
      // describes each parameter.
      if (paramAnnotation.annotationType().equals(OpuxlArg.class)) {
        result.add(new ExcelArgumentAttribute((OpuxlArg) paramAnnotation));
      }
    }
    return result;
  }

  /**
   * Executes the function described by the given Function Payload.
   * @param payload
   *          the payload which describes the function to execute
   * @return the Matrix result of the function
   */
  public OpuxlMatrix execute(final FunctionPayload payload) {
    if (!hasFunction(payload.getName())) {
      throw new NoSuchMethodError("No method found with name: "
          + payload.getName());
    }
    LOG.info("Executing Function: {} with Args: {}", payload.getName(), payload
        .getArgs());
    return getMethods().get(payload.getName())
        .execute(normalizeArgs(payload.getArgs()));
  }

  /**
   * We have to normalize the arguments sent via Excel as Excel sends different
   * formats for Array vs Array-as-Range.
   * <p>
   * Array as Range method(E9:H9) => ["name", "id", "foobar"]
   * <p>
   * Array method({"name";"id";"foobar"}) => [["name"], ["id"], ["foobar"]]
   * @param args
   *          the arguments provided via excel
   * @return the arguments in a normalized format
   */
  @SuppressWarnings({
    "unchecked"
  })
  private List<Object> normalizeArgs(final List<Object> args) {
    final List<Object> result = Lists.newArrayList();

    for (final Object object : args) {
      if (object instanceof List) {
        result.add(flattenList((List<Object>) object));
      } else {
        result.add(object);
      }
    }

    return result;
  }

  @SuppressWarnings("unchecked")
  private List<Object> flattenList(final List<Object> list) {
    final List<Object> result = Lists.newArrayList();

    for (final Object object : list) {
      if (object instanceof List) {
        for (final Object item : ((List<Object>) object)) {
          result.add(item);
        }
      } else {
        result.add(object);
      }
    }

    return result;
  }

  /**
   * Checks, whether a method for the given name is registered.
   * @param name
   *          the method name
   * @return true, if a method exists for the name
   */
  public boolean hasFunction(final String name) {
    return getMethods().containsKey(name);
  }

  public Map<String, InstanceMethod> getMethods() {
    return methods;
  }

}
