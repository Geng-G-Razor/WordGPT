/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * FunctionFile entry kept for Office command compatibility.
 * The current ribbon button opens the task pane directly via manifest.xml,
 * so this function intentionally performs no document mutation.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
        ? global
        : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;
