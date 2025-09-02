package com.automationanywhere.botcommand.excel.commands;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.BotCommand;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.SessionValue;
import com.automationanywhere.botcommand.excel.sessions.WorkbookSession;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import com.automationanywhere.toolchain.runtime.session.CloseableSessionObject;
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

public final class RenameWorksheetCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(RenameWorksheetCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    RenameWorksheet command = new RenameWorksheet();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("selectMode") && parameters.get("selectMode") != null && parameters.get("selectMode").get() != null) {
      convertedParameters.put("selectMode", parameters.get("selectMode").get());
      if(convertedParameters.get("selectMode") !=null && !(convertedParameters.get("selectMode") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","selectMode", "String", parameters.get("selectMode").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("selectMode") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","selectMode"));
    }
    if(convertedParameters.get("selectMode") != null) {
      switch((String)convertedParameters.get("selectMode")) {
        case "ByIndex" : {
          if(parameters.containsKey("originalIndexOneBased") && parameters.get("originalIndexOneBased") != null && parameters.get("originalIndexOneBased").get() != null) {
            convertedParameters.put("originalIndexOneBased", parameters.get("originalIndexOneBased").get());
            if(convertedParameters.get("originalIndexOneBased") !=null && !(convertedParameters.get("originalIndexOneBased") instanceof Double)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","originalIndexOneBased", "Double", parameters.get("originalIndexOneBased").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("originalIndexOneBased") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","originalIndexOneBased"));
          }


        } break;
        case "ByName" : {
          if(parameters.containsKey("originalName") && parameters.get("originalName") != null && parameters.get("originalName").get() != null) {
            convertedParameters.put("originalName", parameters.get("originalName").get());
            if(convertedParameters.get("originalName") !=null && !(convertedParameters.get("originalName") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","originalName", "String", parameters.get("originalName").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("originalName") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","originalName"));
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","selectMode"));
      }
    }

    if(parameters.containsKey("newName") && parameters.get("newName") != null && parameters.get("newName").get() != null) {
      convertedParameters.put("newName", parameters.get("newName").get());
      if(convertedParameters.get("newName") !=null && !(convertedParameters.get("newName") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","newName", "String", parameters.get("newName").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("newName") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","newName"));
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
      command.action((String)convertedParameters.get("selectMode"),(Double)convertedParameters.get("originalIndexOneBased"),(String)convertedParameters.get("originalName"),(String)convertedParameters.get("newName"),(WorkbookSession)convertedParameters.get("session"));Optional<Value> result = Optional.empty();
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
