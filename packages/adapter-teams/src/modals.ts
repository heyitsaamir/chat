/**
 * Teams dialog (task module) converter.
 * Converts ModalElement to Adaptive Card JSON for Teams task modules,
 * and converts ModalResponse to TaskModuleResponse format.
 */

import { createEmojiConverter, mapButtonStyle } from "@chat-adapter/shared";
import type {
  FieldsElement,
  ModalChild,
  ModalElement,
  ModalResponse,
  RadioSelectElement,
  SelectElement,
  TextElement,
  TextInputElement,
} from "chat";

import type {
  AdaptiveCard,
  AdaptiveCardAction,
  AdaptiveCardElement,
} from "./cards";

const convertEmoji = createEmojiConverter("teams");

const ADAPTIVE_CARD_SCHEMA =
  "http://adaptivecards.io/schemas/adaptive-card.json";
const ADAPTIVE_CARD_VERSION = "1.4";

// ============================================================================
// ModalElement -> Adaptive Card
// ============================================================================

/**
 * Convert a ModalElement to an Adaptive Card for use inside a Teams task module.
 *
 * @param modal - The modal element to convert
 * @param contextId - Context ID for server-side stored thread/message context
 * @param callbackId - Callback ID for routing the submit event
 */
export function modalToAdaptiveCard(
  modal: ModalElement,
  contextId: string,
  callbackId: string
): AdaptiveCard {
  const body: AdaptiveCardElement[] = [];

  for (const child of modal.children) {
    body.push(...modalChildToAdaptiveElements(child));
  }

  const submitData: Record<string, unknown> = {
    __contextId: contextId,
    __callbackId: callbackId,
  };

  const submitAction: AdaptiveCardAction = {
    type: "Action.Submit",
    title: modal.submitLabel || "Submit",
    data: submitData,
  };

  const style = mapButtonStyle("primary", "teams");
  if (style) {
    submitAction.style = style;
  }

  return {
    type: "AdaptiveCard",
    $schema: ADAPTIVE_CARD_SCHEMA,
    version: ADAPTIVE_CARD_VERSION,
    body,
    actions: [submitAction],
  };
}

function modalChildToAdaptiveElements(
  child: ModalChild
): AdaptiveCardElement[] {
  switch (child.type) {
    case "text_input":
      return [textInputToAdaptive(child)];
    case "select":
      return [selectToAdaptive(child)];
    case "radio_select":
      return [radioSelectToAdaptive(child)];
    case "text":
      return [textToAdaptive(child)];
    case "fields":
      return [fieldsToAdaptive(child)];
    default:
      return [];
  }
}

function textInputToAdaptive(input: TextInputElement): AdaptiveCardElement {
  const element: AdaptiveCardElement = {
    type: "Input.Text",
    id: input.id,
    label: convertEmoji(input.label),
    isMultiline: input.multiline ?? false,
    isRequired: !(input.optional ?? false),
  };

  if (input.placeholder) {
    element.placeholder = convertEmoji(input.placeholder);
  }
  if (input.initialValue) {
    element.value = input.initialValue;
  }
  if (input.maxLength) {
    element.maxLength = input.maxLength;
  }

  return element;
}

function selectToAdaptive(select: SelectElement): AdaptiveCardElement {
  const choices = select.options.map((opt) => ({
    title: convertEmoji(opt.label),
    value: opt.value,
  }));

  const element: AdaptiveCardElement = {
    type: "Input.ChoiceSet",
    id: select.id,
    label: convertEmoji(select.label),
    style: "compact",
    isRequired: !(select.optional ?? false),
    choices,
  };

  if (select.placeholder) {
    element.placeholder = convertEmoji(select.placeholder);
  }
  if (select.initialOption) {
    element.value = select.initialOption;
  }

  return element;
}

function radioSelectToAdaptive(
  radioSelect: RadioSelectElement
): AdaptiveCardElement {
  const choices = radioSelect.options.map((opt) => ({
    title: convertEmoji(opt.label),
    value: opt.value,
  }));

  const element: AdaptiveCardElement = {
    type: "Input.ChoiceSet",
    id: radioSelect.id,
    label: convertEmoji(radioSelect.label),
    style: "expanded",
    isRequired: !(radioSelect.optional ?? false),
    choices,
  };

  if (radioSelect.initialOption) {
    element.value = radioSelect.initialOption;
  }

  return element;
}

function textToAdaptive(text: TextElement): AdaptiveCardElement {
  const block: AdaptiveCardElement = {
    type: "TextBlock",
    text: convertEmoji(text.content),
    wrap: true,
  };

  if (text.style === "bold") {
    block.weight = "bolder";
  } else if (text.style === "muted") {
    block.isSubtle = true;
  }

  return block;
}

function fieldsToAdaptive(fields: FieldsElement): AdaptiveCardElement {
  const facts = fields.children.map((field) => ({
    title: convertEmoji(field.label),
    value: convertEmoji(field.value),
  }));

  return {
    type: "FactSet",
    facts,
  };
}

// ============================================================================
// Dialog submit value parsing
// ============================================================================

export interface DialogSubmitValues {
  callbackId: string | undefined;
  contextId: string | undefined;
  values: Record<string, string>;
}

/**
 * Extract user input values from an Action.Submit data payload,
 * stripping out internal keys (__contextId, __callbackId, msteams).
 */
export function parseDialogSubmitValues(
  data: Record<string, unknown> | undefined
): DialogSubmitValues {
  if (!data) {
    return { contextId: undefined, callbackId: undefined, values: {} };
  }

  const contextId = data.__contextId as string | undefined;
  const callbackId = data.__callbackId as string | undefined;

  const values: Record<string, string> = {};
  for (const [key, val] of Object.entries(data)) {
    if (key === "__contextId" || key === "__callbackId" || key === "msteams") {
      continue;
    }
    if (typeof val === "string") {
      values[key] = val;
    }
  }

  return { contextId, callbackId, values };
}

// ============================================================================
// ModalResponse -> Teams task module response
// ============================================================================

/**
 * Convert a ModalResponse from the handler into a Teams task module response.
 * Returns undefined to signal "close dialog" (empty HTTP body).
 *
 * @param response - The modal response from the submit handler
 * @param stashedModal - The original modal element (for error re-rendering)
 * @param logger - Optional logger for warnings
 */
export function modalResponseToTaskModuleResponse(
  response: ModalResponse | undefined,
  logger?: { warn: (msg: string, meta?: Record<string, unknown>) => void }
): Record<string, unknown> | undefined {
  if (!response) {
    return undefined;
  }
  switch (response.action) {
    case "close":
      // undefined signals "close dialog" (empty HTTP body)
      return undefined;

    case "update": {
      const card = modalToAdaptiveCard(
        response.modal,
        "", // contextId not needed for update re-render
        response.modal.callbackId
      );
      return {
        task: {
          type: "continue",
          value: {
            title: response.modal.title,
            card: {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: card,
            },
          },
        },
      };
    }

    case "push": {
      // Teams has no dialog stacking — fall back to update with a warning
      logger?.warn(
        "Teams does not support dialog stacking (push). Falling back to update.",
        { callbackId: response.modal.callbackId }
      );
      const card = modalToAdaptiveCard(
        response.modal,
        "",
        response.modal.callbackId
      );
      return {
        task: {
          type: "continue",
          value: {
            title: response.modal.title,
            card: {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: card,
            },
          },
        },
      };
    }

    case "errors": {
      // Render a simple error card listing validation issues
      const errorLines = Object.entries(response.errors).map(
        ([field, msg]) => ({
          type: "TextBlock",
          text: `**${field}**: ${msg}`,
          wrap: true,
          color: "attention",
        })
      );

      const errorCard: AdaptiveCard = {
        type: "AdaptiveCard",
        $schema: ADAPTIVE_CARD_SCHEMA,
        version: ADAPTIVE_CARD_VERSION,
        body: [
          {
            type: "TextBlock",
            text: "Please fix the following errors:",
            weight: "bolder",
            wrap: true,
          },
          ...errorLines,
        ],
      };

      return {
        task: {
          type: "continue",
          value: {
            title: "Validation Error",
            card: {
              contentType: "application/vnd.microsoft.card.adaptive",
              content: errorCard,
            },
          },
        },
      };
    }

    default:
      return undefined;
  }
}
