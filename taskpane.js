
/*
 Retainer & Credit Card Authorization Generator
 Microsoft Word for Mac
*/

const $ = id => document.getElementById(id);

const TEMPLATE_BASE =
  "https://tynerjim2026.github.io/retainer-generator/";

const RETAINER_TEMPLATE =
  TEMPLATE_BASE + "Retainer.docx";

const CARD_TEMPLATE =
  TEMPLATE_BASE + "Card.docx";

function showStatus(message) {
  const status = $("status");
  status.style.display = "block";
  status.textContent = message;
}

function formatMoney(amount) {
  return Number(amount).toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function numberToWords(n) {
  const ones = [
    "", "one", "two", "three", "four", "five",
    "six", "seven", "eight", "nine", "ten",
    "eleven", "twelve", "thirteen", "fourteen",
    "fifteen", "sixteen", "seventeen",
    "eighteen", "nineteen"
  ];

  const tens = [
    "", "", "twenty", "thirty", "forty",
    "fifty", "sixty", "seventy", "eighty", "ninety"
  ];

  function underThousand(value) {
    let words = [];

    if (value >= 100) {
      words.push(ones[Math.floor(value / 100)] + " hundred");
      value %= 100;
    }

    if (value >= 20) {
      let part = tens[Math.floor(value / 10)];
      if (value % 10) {
        part += "-" + ones[value % 10];
      }
      words.push(part);
    } else if (value > 0) {
      words.push(ones[value]);
    }

    return words.join(" ");
  }

  if (n === 0) return "zero";

  const groups = [
    [1000000000, "billion"],
    [1000000, "million"],
    [1000, "thousand"],
    [1, ""]
  ];

  let words = [];

  for (const [size, label] of groups) {
    const group = Math.floor(n / size);

    if (group > 0) {
      let part = underThousand(group);
      if (label) part += " " + label;
      words.push(part);
      n %= size;
    }
  }

  return words.join(" ");
}

function getAmount() {
  const raw = $("retainerAmount").value.trim();

  if (!/^\$?\s*(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d{1,2})?$/.test(raw)) {
    throw new Error(
      "Enter a valid dollar amount, such as 1500 or 1500.00."
    );
  }

  const amount = Number(raw.replace(/[$,\s]/g, ""));

  if (!Number.isFinite(amount) ||
      amount <= 0 ||
      amount >= 1000000000000) {
    throw new Error("Enter a valid positive retainer amount.");
  }

  return amount;
}

function getData(forRetainer) {
  const amount = getAmount();

  const client = $("clientName").value.trim();
  const matter = $("legalMatter").value.trim();

  if (forRetainer && !client) {
    throw new Error("Enter the client's full name.");
  }

  if (forRetainer && !matter) {
    throw new Error("Enter the legal matter.");
  }

  const cents = Math.round(amount * 100);
  const dollars = Math.floor(cents / 100);
  const fraction = cents % 100;

  return {
    client,
    matter,
    amount: formatMoney(amount),
    amountWords:
      numberToWords(dollars) +
      " dollars and " +
      String(fraction).padStart(2, "0") +
      "/100"
  };
}

/*
 Search and replace a placeholder in the
 currently open Word document.
*/
async function replacePlaceholder(context, marker, value) {
  const matches = context.document.body.search(marker, {
    matchCase: true,
    matchWholeWord: false
  });

  matches.load("items");
  await context.sync();

  if (matches.items.length === 0) {
    throw new Error("Missing template placeholder: " + marker);
  }

  for (const match of matches.items) {
    match.insertText(
      String(value),
      Word.InsertLocation.replace
    );
  }

  await context.sync();
}

/*
 Insert a fresh template as a new document.
*/
async function openTemplate(url) {
  if (!Office.context.requirements.isSetSupported(
    "WordApi",
    "1.3"
  )) {
    throw new Error(
      "This Word version does not support the required document API."
    );
  }

  const response = await fetch(url, {
    cache: "no-store"
  });

  if (!response.ok) {
    throw new Error(
      "Could not load Word template. HTTP " + response.status
    );
  }

  const bytes = new Uint8Array(await response.arrayBuffer());

  let binary = "";
  const chunkSize = 8192;

  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(
      null,
      bytes.subarray(i, i + chunkSize)
    );
  }

  const base64 = btoa(binary);

  await Word.run(async context => {
    context.application.createDocument(base64).open();
    await context.sync();
  });
}

/*
 Fill the current document.
 The template must contain the expected markers.
*/
async function fillRetainer(data) {
  await Word.run(async context => {
    const fields = [
      ["{{CLIENT_NAME}}", data.client],
      ["{{LEGAL_MATTER}}", data.matter],
      ["{{RETAINER_AMOUNT}}", data.amount],
      ["{{RETAINER_AMOUNT_WORDS}}", data.amountWords]
    ];

    for (const [marker, value] of fields) {
      await replacePlaceholder(context, marker, value);
    }
  });
}

async function fillCard(data) {
  await Word.run(async context => {
    await replacePlaceholder(
      context,
      "{{CHARGE_AMOUNT}}",
      data.amount
    );
  });
}

/*
 Generate the chosen document.
*/
async function generate(type) {
  const button = $(
    type === "retainer" ? "generateRetainer" : "generateCard"
  );

  button.disabled = true;

  try {
    const data = getData(type === "retainer");

    const url = type === "retainer"
      ? RETAINER_TEMPLATE
      : CARD_TEMPLATE;

    showStatus("Opening " +
      (type === "retainer"
        ? "Retainer Agreement"
        : "Credit Card Authorization") +
      " template...");

    await openTemplate(url);

    /*
      Word for Mac may switch to the new document
      asynchronously. Confirm the active document
      before filling it.
    */
    showStatus(
      "The template has been opened in Word. " +
      "Select the newly opened document, then " +
      "click the same Generate button again " +
      "to fill its fields."
    );

  } catch (error) {
    showStatus(
      "Unable to generate document: " +
      (error.message || String(error))
    );

    console.error(error);

  } finally {
    button.disabled = false;
  }
}

/*
 Start the questionnaire.
*/
Office.onReady(() => {
  $("generateRetainer").addEventListener("click", () => {
    generate("retainer");
  });

  $("generateCard").addEventListener("click", () => {
    generate("card");
  });

  showStatus(
    "Enter the information, then select the document to generate."
  );
});
