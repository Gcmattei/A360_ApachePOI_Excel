package com.davita.botcommand.excel.commands;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.BotCommand;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import java.lang.Boolean;
import java.lang.ClassCastException;
import java.lang.Deprecated;
import java.lang.Object;
import java.lang.String;
import java.lang.Throwable;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public final class OpenWorkbookCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(OpenWorkbookCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    OpenWorkbook command = new OpenWorkbook();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("filePath") && parameters.get("filePath") != null && parameters.get("filePath").get() != null) {
      convertedParameters.put("filePath", parameters.get("filePath").get());
      if(convertedParameters.get("filePath") !=null && !(convertedParameters.get("filePath") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","filePath", "String", parameters.get("filePath").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("filePath") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","filePath"));
    }
    if(convertedParameters.containsKey("filePath")) {
      String filePath= ((String)convertedParameters.get("filePath"));
      int lastIndxDot = filePath.lastIndexOf(".");
      if (lastIndxDot == -1 || lastIndxDot >= filePath.length()) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.FileExtension","filePath","xlsx,xls"));
      }
      String fileExtension = filePath.substring(lastIndxDot + 1);
      if(!Arrays.stream("xlsx,xls".split(",")).anyMatch(fileExtension::equalsIgnoreCase))  {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.FileExtension","filePath","xlsx,xls"));
      }

    }
    if(parameters.containsKey("readOnly") && parameters.get("readOnly") != null && parameters.get("readOnly").get() != null) {
      convertedParameters.put("readOnly", parameters.get("readOnly").get());
      if(convertedParameters.get("readOnly") !=null && !(convertedParameters.get("readOnly") instanceof Boolean)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","readOnly", "Boolean", parameters.get("readOnly").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("readOnly") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","readOnly"));
    }

    try {
      Optional<Value> result =  Optional.ofNullable(command.action((String)convertedParameters.get("filePath"),(Boolean)convertedParameters.get("readOnly")));
      return logger.traceExit(result);
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.IllegalParameters","action"));
    }
    catch (BotCommandException e) {
      logger.fatal(e.getMessage(),e);
      throw e;
    }
    catch (Throwable e) {
      logger.fatal(e.getMessage(),e);
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.NotBotCommandException",e.getMessage()),e);
    }
  }

  public Map<String, Value> executeAndReturnMany(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return null;
  }
}
