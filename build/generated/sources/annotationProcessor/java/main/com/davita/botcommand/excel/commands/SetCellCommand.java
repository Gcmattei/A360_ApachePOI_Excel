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

public final class SetCellCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(SetCellCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    SetCell command = new SetCell();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("targetMode") && parameters.get("targetMode") != null && parameters.get("targetMode").get() != null) {
      convertedParameters.put("targetMode", parameters.get("targetMode").get());
      if(convertedParameters.get("targetMode") !=null && !(convertedParameters.get("targetMode") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","targetMode", "String", parameters.get("targetMode").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("targetMode") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","targetMode"));
    }
    if(convertedParameters.get("targetMode") != null) {
      switch((String)convertedParameters.get("targetMode")) {
        case "ACTIVE" : {

        } break;
        case "SPECIFIC" : {
          if(parameters.containsKey("cellOrRangeA1") && parameters.get("cellOrRangeA1") != null && parameters.get("cellOrRangeA1").get() != null) {
            convertedParameters.put("cellOrRangeA1", parameters.get("cellOrRangeA1").get());
            if(convertedParameters.get("cellOrRangeA1") !=null && !(convertedParameters.get("cellOrRangeA1") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","cellOrRangeA1", "String", parameters.get("cellOrRangeA1").get().getClass().getSimpleName()));
            }
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","targetMode"));
      }
    }

    if(parameters.containsKey("input") && parameters.get("input") != null && parameters.get("input").get() != null) {
      convertedParameters.put("input", parameters.get("input").get());
      if(convertedParameters.get("input") !=null && !(convertedParameters.get("input") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","input", "String", parameters.get("input").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("input") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","input"));
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
      command.action((String)convertedParameters.get("targetMode"),(String)convertedParameters.get("cellOrRangeA1"),(String)convertedParameters.get("input"),(WorkbookSession)convertedParameters.get("session"));Optional<Value> result = Optional.empty();
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
