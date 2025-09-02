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

public final class FilterWorksheetOrTableCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(FilterWorksheetOrTableCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.entrySet().stream().filter(en -> !Arrays.asList( new String[] {}).contains(en.getKey()) && en.getValue() != null).collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue)).toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    FilterWorksheetOrTable command = new FilterWorksheetOrTable();
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
        case "TABLE" : {
          if(parameters.containsKey("tableName") && parameters.get("tableName") != null && parameters.get("tableName").get() != null) {
            convertedParameters.put("tableName", parameters.get("tableName").get());
            if(convertedParameters.get("tableName") !=null && !(convertedParameters.get("tableName") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableName", "String", parameters.get("tableName").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("tableName") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","tableName"));
          }

          if(parameters.containsKey("tableColSelector") && parameters.get("tableColSelector") != null && parameters.get("tableColSelector").get() != null) {
            convertedParameters.put("tableColSelector", parameters.get("tableColSelector").get());
            if(convertedParameters.get("tableColSelector") !=null && !(convertedParameters.get("tableColSelector") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableColSelector", "String", parameters.get("tableColSelector").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("tableColSelector") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","tableColSelector"));
          }
          if(convertedParameters.get("tableColSelector") != null) {
            switch((String)convertedParameters.get("tableColSelector")) {
              case "BY_NAME" : {
                if(parameters.containsKey("tableColumnName") && parameters.get("tableColumnName") != null && parameters.get("tableColumnName").get() != null) {
                  convertedParameters.put("tableColumnName", parameters.get("tableColumnName").get());
                  if(convertedParameters.get("tableColumnName") !=null && !(convertedParameters.get("tableColumnName") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableColumnName", "String", parameters.get("tableColumnName").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("tableColumnName") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","tableColumnName"));
                }


              } break;
              case "BY_INDEX" : {
                if(parameters.containsKey("tableColumnIndexOneBased") && parameters.get("tableColumnIndexOneBased") != null && parameters.get("tableColumnIndexOneBased").get() != null) {
                  convertedParameters.put("tableColumnIndexOneBased", parameters.get("tableColumnIndexOneBased").get());
                  if(convertedParameters.get("tableColumnIndexOneBased") !=null && !(convertedParameters.get("tableColumnIndexOneBased") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","tableColumnIndexOneBased", "Double", parameters.get("tableColumnIndexOneBased").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("tableColumnIndexOneBased") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","tableColumnIndexOneBased"));
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","tableColSelector"));
            }
          }


        } break;
        case "WORKSHEET" : {
          if(parameters.containsKey("sheetName") && parameters.get("sheetName") != null && parameters.get("sheetName").get() != null) {
            convertedParameters.put("sheetName", parameters.get("sheetName").get());
            if(convertedParameters.get("sheetName") !=null && !(convertedParameters.get("sheetName") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","sheetName", "String", parameters.get("sheetName").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("sheetName") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","sheetName"));
          }

          if(parameters.containsKey("wsRangeMode") && parameters.get("wsRangeMode") != null && parameters.get("wsRangeMode").get() != null) {
            convertedParameters.put("wsRangeMode", parameters.get("wsRangeMode").get());
            if(convertedParameters.get("wsRangeMode") !=null && !(convertedParameters.get("wsRangeMode") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsRangeMode", "String", parameters.get("wsRangeMode").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("wsRangeMode") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsRangeMode"));
          }
          if(convertedParameters.get("wsRangeMode") != null) {
            switch((String)convertedParameters.get("wsRangeMode")) {
              case "ALL" : {

              } break;
              case "SPECIFIC" : {
                if(parameters.containsKey("wsRangeA1") && parameters.get("wsRangeA1") != null && parameters.get("wsRangeA1").get() != null) {
                  convertedParameters.put("wsRangeA1", parameters.get("wsRangeA1").get());
                  if(convertedParameters.get("wsRangeA1") !=null && !(convertedParameters.get("wsRangeA1") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsRangeA1", "String", parameters.get("wsRangeA1").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("wsRangeA1") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsRangeA1"));
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","wsRangeMode"));
            }
          }

          if(parameters.containsKey("wsColSelector") && parameters.get("wsColSelector") != null && parameters.get("wsColSelector").get() != null) {
            convertedParameters.put("wsColSelector", parameters.get("wsColSelector").get());
            if(convertedParameters.get("wsColSelector") !=null && !(convertedParameters.get("wsColSelector") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsColSelector", "String", parameters.get("wsColSelector").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("wsColSelector") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsColSelector"));
          }
          if(convertedParameters.get("wsColSelector") != null) {
            switch((String)convertedParameters.get("wsColSelector")) {
              case "BY_NAME" : {
                if(parameters.containsKey("wsColumnName") && parameters.get("wsColumnName") != null && parameters.get("wsColumnName").get() != null) {
                  convertedParameters.put("wsColumnName", parameters.get("wsColumnName").get());
                  if(convertedParameters.get("wsColumnName") !=null && !(convertedParameters.get("wsColumnName") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsColumnName", "String", parameters.get("wsColumnName").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("wsColumnName") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsColumnName"));
                }


              } break;
              case "BY_INDEX" : {
                if(parameters.containsKey("wsColumnIndexOneBased") && parameters.get("wsColumnIndexOneBased") != null && parameters.get("wsColumnIndexOneBased").get() != null) {
                  convertedParameters.put("wsColumnIndexOneBased", parameters.get("wsColumnIndexOneBased").get());
                  if(convertedParameters.get("wsColumnIndexOneBased") !=null && !(convertedParameters.get("wsColumnIndexOneBased") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","wsColumnIndexOneBased", "Double", parameters.get("wsColumnIndexOneBased").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("wsColumnIndexOneBased") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","wsColumnIndexOneBased"));
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","wsColSelector"));
            }
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","mode"));
      }
    }

    if(parameters.containsKey("filterType") && parameters.get("filterType") != null && parameters.get("filterType").get() != null) {
      convertedParameters.put("filterType", parameters.get("filterType").get());
      if(convertedParameters.get("filterType") !=null && !(convertedParameters.get("filterType") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","filterType", "String", parameters.get("filterType").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("filterType") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","filterType"));
    }
    if(convertedParameters.get("filterType") != null) {
      switch((String)convertedParameters.get("filterType")) {
        case "NUMBER" : {
          if(parameters.containsKey("numOperator") && parameters.get("numOperator") != null && parameters.get("numOperator").get() != null) {
            convertedParameters.put("numOperator", parameters.get("numOperator").get());
            if(convertedParameters.get("numOperator") !=null && !(convertedParameters.get("numOperator") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","numOperator", "String", parameters.get("numOperator").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("numOperator") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","numOperator"));
          }
          if(convertedParameters.get("numOperator") != null) {
            switch((String)convertedParameters.get("numOperator")) {
              case "EQ" : {
                if(parameters.containsKey("numValueEq") && parameters.get("numValueEq") != null && parameters.get("numValueEq").get() != null) {
                  convertedParameters.put("numValueEq", parameters.get("numValueEq").get());
                  if(convertedParameters.get("numValueEq") !=null && !(convertedParameters.get("numValueEq") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","numValueEq", "Double", parameters.get("numValueEq").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("numValueEq") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","numValueEq"));
                }


              } break;
              case "NEQ" : {
                if(parameters.containsKey("numValueNeq") && parameters.get("numValueNeq") != null && parameters.get("numValueNeq").get() != null) {
                  convertedParameters.put("numValueNeq", parameters.get("numValueNeq").get());
                  if(convertedParameters.get("numValueNeq") !=null && !(convertedParameters.get("numValueNeq") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","numValueNeq", "Double", parameters.get("numValueNeq").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("numValueNeq") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","numValueNeq"));
                }


              } break;
              case "GT" : {
                if(parameters.containsKey("numValueGt") && parameters.get("numValueGt") != null && parameters.get("numValueGt").get() != null) {
                  convertedParameters.put("numValueGt", parameters.get("numValueGt").get());
                  if(convertedParameters.get("numValueGt") !=null && !(convertedParameters.get("numValueGt") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","numValueGt", "Double", parameters.get("numValueGt").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("numValueGt") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","numValueGt"));
                }


              } break;
              case "GTE" : {
                if(parameters.containsKey("numValueGte") && parameters.get("numValueGte") != null && parameters.get("numValueGte").get() != null) {
                  convertedParameters.put("numValueGte", parameters.get("numValueGte").get());
                  if(convertedParameters.get("numValueGte") !=null && !(convertedParameters.get("numValueGte") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","numValueGte", "Double", parameters.get("numValueGte").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("numValueGte") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","numValueGte"));
                }


              } break;
              case "LT" : {
                if(parameters.containsKey("numValueLt") && parameters.get("numValueLt") != null && parameters.get("numValueLt").get() != null) {
                  convertedParameters.put("numValueLt", parameters.get("numValueLt").get());
                  if(convertedParameters.get("numValueLt") !=null && !(convertedParameters.get("numValueLt") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","numValueLt", "Double", parameters.get("numValueLt").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("numValueLt") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","numValueLt"));
                }


              } break;
              case "LTE" : {
                if(parameters.containsKey("numValueLte") && parameters.get("numValueLte") != null && parameters.get("numValueLte").get() != null) {
                  convertedParameters.put("numValueLte", parameters.get("numValueLte").get());
                  if(convertedParameters.get("numValueLte") !=null && !(convertedParameters.get("numValueLte") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","numValueLte", "Double", parameters.get("numValueLte").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("numValueLte") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","numValueLte"));
                }


              } break;
              case "BETWEEN" : {
                if(parameters.containsKey("numValueBetween") && parameters.get("numValueBetween") != null && parameters.get("numValueBetween").get() != null) {
                  convertedParameters.put("numValueBetween", parameters.get("numValueBetween").get());
                  if(convertedParameters.get("numValueBetween") !=null && !(convertedParameters.get("numValueBetween") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","numValueBetween", "Double", parameters.get("numValueBetween").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("numValueBetween") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","numValueBetween"));
                }

                if(parameters.containsKey("numValueBetween2") && parameters.get("numValueBetween2") != null && parameters.get("numValueBetween2").get() != null) {
                  convertedParameters.put("numValueBetween2", parameters.get("numValueBetween2").get());
                  if(convertedParameters.get("numValueBetween2") !=null && !(convertedParameters.get("numValueBetween2") instanceof Double)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","numValueBetween2", "Double", parameters.get("numValueBetween2").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("numValueBetween2") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","numValueBetween2"));
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","numOperator"));
            }
          }


        } break;
        case "TEXT" : {
          if(parameters.containsKey("textOperator") && parameters.get("textOperator") != null && parameters.get("textOperator").get() != null) {
            convertedParameters.put("textOperator", parameters.get("textOperator").get());
            if(convertedParameters.get("textOperator") !=null && !(convertedParameters.get("textOperator") instanceof String)) {
              throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","textOperator", "String", parameters.get("textOperator").get().getClass().getSimpleName()));
            }
          }
          if(convertedParameters.get("textOperator") == null) {
            throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","textOperator"));
          }
          if(convertedParameters.get("textOperator") != null) {
            switch((String)convertedParameters.get("textOperator")) {
              case "EQ" : {
                if(parameters.containsKey("textValueEq") && parameters.get("textValueEq") != null && parameters.get("textValueEq").get() != null) {
                  convertedParameters.put("textValueEq", parameters.get("textValueEq").get());
                  if(convertedParameters.get("textValueEq") !=null && !(convertedParameters.get("textValueEq") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","textValueEq", "String", parameters.get("textValueEq").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("textValueEq") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","textValueEq"));
                }


              } break;
              case "NEQ" : {
                if(parameters.containsKey("textValueNeq") && parameters.get("textValueNeq") != null && parameters.get("textValueNeq").get() != null) {
                  convertedParameters.put("textValueNeq", parameters.get("textValueNeq").get());
                  if(convertedParameters.get("textValueNeq") !=null && !(convertedParameters.get("textValueNeq") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","textValueNeq", "String", parameters.get("textValueNeq").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("textValueNeq") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","textValueNeq"));
                }


              } break;
              case "BEG" : {
                if(parameters.containsKey("textValueBeg") && parameters.get("textValueBeg") != null && parameters.get("textValueBeg").get() != null) {
                  convertedParameters.put("textValueBeg", parameters.get("textValueBeg").get());
                  if(convertedParameters.get("textValueBeg") !=null && !(convertedParameters.get("textValueBeg") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","textValueBeg", "String", parameters.get("textValueBeg").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("textValueBeg") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","textValueBeg"));
                }


              } break;
              case "END" : {
                if(parameters.containsKey("textValueEnd") && parameters.get("textValueEnd") != null && parameters.get("textValueEnd").get() != null) {
                  convertedParameters.put("textValueEnd", parameters.get("textValueEnd").get());
                  if(convertedParameters.get("textValueEnd") !=null && !(convertedParameters.get("textValueEnd") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","textValueEnd", "String", parameters.get("textValueEnd").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("textValueEnd") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","textValueEnd"));
                }


              } break;
              case "CON" : {
                if(parameters.containsKey("textValueCon") && parameters.get("textValueCon") != null && parameters.get("textValueCon").get() != null) {
                  convertedParameters.put("textValueCon", parameters.get("textValueCon").get());
                  if(convertedParameters.get("textValueCon") !=null && !(convertedParameters.get("textValueCon") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","textValueCon", "String", parameters.get("textValueCon").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("textValueCon") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","textValueCon"));
                }


              } break;
              case "NCON" : {
                if(parameters.containsKey("textValueNcon") && parameters.get("textValueNcon") != null && parameters.get("textValueNcon").get() != null) {
                  convertedParameters.put("textValueNcon", parameters.get("textValueNcon").get());
                  if(convertedParameters.get("textValueNcon") !=null && !(convertedParameters.get("textValueNcon") instanceof String)) {
                    throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","textValueNcon", "String", parameters.get("textValueNcon").get().getClass().getSimpleName()));
                  }
                }
                if(convertedParameters.get("textValueNcon") == null) {
                  throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","textValueNcon"));
                }


              } break;
              default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","textOperator"));
            }
          }


        } break;
        default : throw new BotCommandException(MESSAGES_GENERIC.getString("generic.InvalidOption","filterType"));
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
      command.action((String)convertedParameters.get("mode"),(String)convertedParameters.get("tableName"),(String)convertedParameters.get("tableColSelector"),(String)convertedParameters.get("tableColumnName"),(Double)convertedParameters.get("tableColumnIndexOneBased"),(String)convertedParameters.get("sheetName"),(String)convertedParameters.get("wsRangeMode"),(String)convertedParameters.get("wsRangeA1"),(String)convertedParameters.get("wsColSelector"),(String)convertedParameters.get("wsColumnName"),(Double)convertedParameters.get("wsColumnIndexOneBased"),(String)convertedParameters.get("filterType"),(String)convertedParameters.get("numOperator"),(Double)convertedParameters.get("numValueEq"),(Double)convertedParameters.get("numValueNeq"),(Double)convertedParameters.get("numValueGt"),(Double)convertedParameters.get("numValueGte"),(Double)convertedParameters.get("numValueLt"),(Double)convertedParameters.get("numValueLte"),(Double)convertedParameters.get("numValueBetween"),(Double)convertedParameters.get("numValueBetween2"),(String)convertedParameters.get("textOperator"),(String)convertedParameters.get("textValueEq"),(String)convertedParameters.get("textValueNeq"),(String)convertedParameters.get("textValueBeg"),(String)convertedParameters.get("textValueEnd"),(String)convertedParameters.get("textValueCon"),(String)convertedParameters.get("textValueNcon"),(WorkbookSession)convertedParameters.get("session"));Optional<Value> result = Optional.empty();
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
