
/*
  Retainer Agreement Generator
  Microsoft Word for Mac
  Version 4 - Single Generate Button
*/

const $ = id => document.getElementById(id);

function showStatus(message) {
  const el = $("status");
  if (el) {
    el.style.display = "block";
    el.textContent = message;
  }
}

function amountDetails() {
  const raw = $("retainerAmount").value.trim();

  if (!/^(?:\$)?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d{1,2})?$/.test(raw)) {
    throw new Error("Enter an amount such as 1500 or 2,500.00.");
  }

  const normalized = raw.replace(/[$,]/g, "");
  const [whole, decimal = ""] = normalized.split(".");
  const dollars = Number(whole);
  const cents = Number(decimal.padEnd(2, "0"));

  if (
    !Number.isSafeInteger(dollars) ||
    dollars <= 0 ||
    dollars >= 1000000000000
  ) {
    throw new Error(
      "Enter a positive amount below one trillion dollars."
    );
  }

  return {
    dollars,
    cents,
    formatted:
      "$" +
      dollars.toLocaleString("en-US") +
      "." +
      String(cents).padStart(2, "0")
  };
}

function numberWords(number) {
  const small = [
    "zero", "one", "two", "three", "four", "five",
    "six", "seven", "eight", "nine", "ten",
    "eleven", "twelve", "thirteen", "fourteen",
    "fifteen", "sixteen", "seventeen",
    "eighteen", "nineteen"
  ];

  const tens = [
    "", "", "twenty", "thirty", "forty",
    "fifty", "sixty", "seventy", "eighty", "ninety"
  ];

  function belowThousand(n) {
    if (n < 20) return small[n];

    if (n < 100) {
      return (
        tens[Math.floor(n / 10)] +
        (n % 10 ? "-" + small[n % 10] : "")
      );
    }

    return (
      small[Math.floor(n / 100)] +
      " hundred" +
      (n % 100 ? " " + belowThousand(n % 100) : "")
    );
  }

  if (number === 0) return "zero";

  const scales = [
    [1000000000, "billion"],
    [1000000, "million"],
    [1000, "thousand"],
    [1, ""]
  ];

  const parts = [];

  for (const [size, label] of scales) {
    const group = Math.floor(number / size);

    if (group) {
      parts.push(
        belowThousand(group) + (label ? " " + label : "")
      );
      number %= size;
    }
  }

  return parts.join(" ");
}

function getData() {
  const amount = amountDetails();

  const client = $("clientName").value.trim();
  const matter = $("legalMatter").value.trim();

  if (!client) {
    throw new Error("Enter the client's full name.");
  }

  if (!matter) {
    throw new Error("Enter the legal matter.");
  }

  if (client.length > 200 || matter.length > 1000) {
    throw new Error(
      "Please shorten the name or matter description."
    );
  }

  if (/[\r\n]/.test(client)) {
    throw new Error(
      "The client's name must be on one line."
    );
  }

  return {
    client,
    matter,
    amount: amount.formatted,
    words: numberWords(amount.dollars)
  };
}

function includeCardForm() {
  const selected = document.querySelector(
    'input[name="includeCard"]:checked'
  );

  if (!selected) {
    throw new Error(
      "Select Yes or No for the Credit Card Authorization Form."
    );
  }

  return selected.value === "yes";
}

function requiredFields(data) {
  return [
    ["{{CLIENT_NAME}}", data.client],
    ["{{LEGAL_MATTER}}", data.matter],
    ["{{RETAINER_WORDS}}", data.words],
    ["{{RETAINER_AMOUNT}}", data.amount]
  ];
}

async function searchMatches(context, marker) {
  const results = context.document.body.search(marker, {
    matchCase: true,
    matchWholeWord: false
  });

  results.load("items");
  await context.sync();

  return results.items;
}

async function inspectDocument() {
  return Word.run(async context => {
    const required = [
      "{{CLIENT_NAME}}",
      "{{LEGAL_MATTER}}",
      "{{RETAINER_WORDS}}",
      "{{RETAINER_AMOUNT}}",
      "[[CARD_SECTION]]",
      "[[CARD_START]]",
      "[[CARD_END]]"
    ];

    const results = required.map(marker => {
      const matches = context.document.body.search(marker, {
        matchCase: true,
        matchWholeWord: false
      });

      matches.load("items");
      return matches;
    });

    await context.sync();

    const missing = required.filter(
      (_, index) => results[index].items.length === 0
    );

    if (missing.length) {
      throw new Error(
        "Open a fresh copy of R4Search.docx. " +
        "Missing: " + missing.join(", ")
      );
    }
  });
}

async function replaceAll(context, marker, value) {
  const matches = await searchMatches(context, marker);

  if (!matches.length) {
    throw new Error("Missing field: " + marker);
  }

  for (const match of matches) {
    match.insertText(
      String(value),
      Word.InsertLocation.replace
    );
  }

  await context.sync();
}

async function removeMarkers(context) {
  const markers = [
    "[[CARD_SECTION]]",
    "[[CARD_START]]",
    "[[CARD_END]]"
  ];

  for (const marker of markers) {
    await replaceAll(context, marker, "");
  }
}

async function removeCardSection(context) {
  const matches = await searchMatches(
    context,
    "[[CARD_SECTION]]"
  );

  if (matches.length !== 1) {
    throw new Error(
      "Could not uniquely identify the Credit Card Authorization section."
    );
  }

  // Remove everything from the page-3 marker
  // through the end of the document.

  const start = matches[0];
  const end = context.document.body.getRange("End");

  const sectionToRemove = start.expandTo(end);
  sectionToRemove.delete();

  await context.sync();
}

async function generate() {
  const retainerButton = $("generateRetainer");

  if (!retainerButton) {
    showStatus("Generate button not found in the questionnaire.");
    return;
  }

  retainerButton.disabled = true;

  try {
    const data = getData();
    const includeCard = includeCardForm();

    showStatus("Checking the open Word document...");

    await inspectDocument();

    await Word.run(async context => {
      const fields = requiredFields(data);

      for (const [marker, value] of fields) {
        await replaceAll(context, marker, value);
      }

      if (includeCard) {
        await removeMarkers(context);
      } else {
        await removeCardSection(context);
      }
    });

    showStatus(
      "Document completed successfully. " +
      "Review the page count and formatting, " +
      "then save it under a new filename."
    );

  } catch (error) {
    showStatus(
      "Unable to generate document: " +
      (error.message || String(error))
    );
    console.error(error);

  } finally {
    retainerButton.disabled = false;
  }
}

Office.onReady(() => {
  const retainerButton = $("generateRetainer");

  if (retainerButton) {
    retainerButton.addEventListener(
      "click",
      () => generate()
    );

    showStatus(
      "Open a fresh copy of R4Search.docx, " +
      "enter the information, select whether to include " +
      "the Credit Card Authorization Form, and click Generate."
    );
  } else {
    showStatus(
      "Unable to find the Generate Retainer Agreement button."
    );
  }
});
