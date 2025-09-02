package com.davita.botcommand.excel.commands;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.BotCommand;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.SessionValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import com.automationanywhere.toolchain.runtime.session.CloseableSessionObject;
import com.davita.botcommand.excel.sessions.WorkbookSession;
import java.lang.ClassCastException;
import java.lang.Deprecated;
import java.lang.Double;
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

public final class SwitchToSheetCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(SwitchToSheetCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    SwitchToSheet command = new SwitchToSheet();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("mode") && parameters.get("mode") != null && parameters.get("mode").get() != null) {
      convertedParameters.put("mode", parameters.get("mode").get());
      if(convertedParameters.get("mode") !=null && !(convertedParameters.get("mode") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","mode", "String", parameters.get("mode").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("mode") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","mode"));
    }
    if(convertedParameters.get("mode") != null) {
      switch((String)convertedParameters.get("mode")) {
        case "BY_INDEX" : {
          if(parameters.containsKey("sheetIndexOneBased") && parameters.get("sheetIndexOneBased") != null && parameters.get("sheetIndexOneBased").get() != null) {
            convertedParameters.put("sheetIndexOneBased", parameters.get("sheetIndexOneBased").get());
            if(convertedParameters.get("sheetIndexOneBased") !=null && !(convertedParameters.get("sheetIndexOneBased") instanceof Double)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetIndexOneBased", "Double", parameters.get("sheetIndexOneBased").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("sheetIndexOneBased") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","sheetIndexOneBased"));
          }


        } break;
        case "BY_NAME" : {
          if(parameters.containsKey("sheetName") && parameters.get("sheetName") != null && parameters.get("sheetName").get() != null) {
            convertedParameters.put("sheetName", parameters.get("sheetName").get());
            if(convertedParameters.get("sheetName") !=null && !(convertedParameters.get("sheetName") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetName", "String", parameters.get("sheetName").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("sheetName") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","sheetName"));
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","mode"));
      }
    }

    if(parameters.containsKey("session") && parameters.get("session") != null) {
      if(((SessionValue)parameters.get("session")).hasObjectValue()) {
        convertedParameters.put("session", ((SessionValue)parameters.get("session")).getSession());
      }
      else if(parameters.get("session").get() != null) {
        convertedParameters.put("session", sessionMap.get(parameters.get("session").get()));
      }
      if(convertedParameters.get("session")!=null && ((CloseableSessionObject)convertedParameters.get("session")).isClosed()) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("session.IsAlreadyClosed",parameters.get("session").get()));
      }
      if(convertedParameters.get("session") !=null && !(convertedParameters.get("session") instanceof WorkbookSession)) {
        Class[] interfaces = convertedParameters.get("session").getClass().getInterfaces();
        for (Class iface : interfaces) {
          if(iface.getName().equals(WorkbookSession.class.getCanonicalName())) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("session.package.mismatch","session" ));
          }
        }
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","session", "WorkbookSession", convertedParameters.get("session").getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("session") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","session"));
    }

    try {
      command.action((String)convertedParameters.get("mode"),(Double)convertedParameters.get("sheetIndexOneBased"),(String)convertedParameters.get("sheetName"),(WorkbookSession)convertedParameters.get("session"));Optional<Value> result = Optional.empty();
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
