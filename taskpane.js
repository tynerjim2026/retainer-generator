
/*
  Retainer & Credit Card Authorization Generator
  Microsoft Word for Mac
  Version 2
*/

const $ = id => document.getElementById(id);

function showStatus(message) {
  const el = $("status");
  el.style.display = "block";
  el.textContent = message;
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

  if (!Number.isSafeInteger(dollars) ||
      dollars <= 0 ||
      dollars >= 1000000000000) {
    throw new Error("Enter a positive amount below one trillion dollars.");
  }

  return {
    dollars,
    cents,
    formatted: "$" + dollars.toLocaleString("en-US") +
      "." + String(cents).padStart(2, "0")
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
      return tens[Math.floor(n / 10)] +
        (n % 10 ? "-" + small[n % 10] : "");
    }

    return small[Math.floor(n / 100)] + " hundred" +
      (n % 100 ? " " + belowThousand(n % 100) : "");
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

function getData(type) {
  const amount = amountDetails();

  const client = $("clientName").value.trim();
  const matter = $("legalMatter").value.trim();

  if (type === "retainer" && !client) {
    throw new Error("Enter the client's full name.");
  }

  if (type === "retainer" && !matter) {
    throw new Error("Enter the legal matter.");
  }

  if (client.length > 200 || matter.length > 1000) {
    throw new Error("Please shorten the name or matter description.");
  }

  if (/[\r\n]/.test(client)) {
    throw new Error("The client's name must be on one line.");
  }

  return {
    client,
    matter,
    amount: amount.formatted,
    words: numberWords(amount.dollars)
  };
}

function requiredFields(type, data) {
  if (type === "retainer") {
    return [
      ["{{CLIENT_NAME}}", data.client],
      ["{{LEGAL_MATTER}}", data.matter],
      ["{{RETAINER_WORDS}}", data.words],
      ["{{RETAINER_AMOUNT}}", data.amount]
    ];
  }

  return [
    ["{{RETAINER_AMOUNT}}", data.amount]
  ];
}

async function inspectDocument(type) {
  return Word.run(async context => {
    const required = type === "retainer"
      ? [
          "{{CLIENT_NAME}}",
          "{{LEGAL_MATTER}}",
          "{{RETAINER_WORDS}}",
          "{{RETAINER_AMOUNT}}"
        ]
      : ["{{RETAINER_AMOUNT}}"];

    const results = required.map(marker => {
      const matches = context.document.body.search(marker, {
        matchCase: true,
        matchWholeWord: false
      });

      matches.load("items");
      return matches;
    });

    const otherMarker = type === "card"
      ? context.document.body.search("{{CLIENT_NAME}}", {
          matchCase: true,
          matchWholeWord: false
        })
      : null;

    if (otherMarker) otherMarker.load("items");

    await context.sync();

    const missing = required.filter(
      (_, index) => results[index].items.length === 0
    );

    if (missing.length) {
      throw new Error(
        "The open Word document is not the correct fresh template. " +
        "Missing: " + missing.join(", ")
      );
    }

    if (otherMarker && otherMarker.items.length) {
      throw new Error(
        "This looks like the Retainer Agreement. " +
        "Open the Credit Card Authorization template instead."
      );
    }
  });
}

async function replaceAll(context, marker, value) {
  const results = context.document.body.search(marker, {
    matchCase: true,
    matchWholeWord: false
  });

  results.load("items");
  await context.sync();

  if (!results.items.length) {
    throw new Error("Missing field: " + marker);
  }

  for (const match of results.items) {
    match.insertText(
      String(value),
      Word.InsertLocation.replace
    );
  }

  await context.sync();
}

async function generate(type) {
  const retainerButton = $("generateRetainer");
  const cardButton = $("generateCard");

  retainerButton.disabled = true;
  cardButton.disabled = true;

  try {
    const data = getData(type);

    showStatus("Checking the open Word document...");

    await inspectDocument(type);

    await Word.run(async context => {
      const fields = requiredFields(type, data);

      for (const [marker, value] of fields) {
        await replaceAll(context, marker, value);
      }
    });

    showStatus(
      "Document completed successfully. " +
      "Review it, then save it under a new filename."
    );

  } catch (error) {
    showStatus(
      "Unable to generate document: " +
      (error.message || String(error))
    );
    console.error(error);

  } finally {
    retainerButton.disabled = false;
    cardButton.disabled = false;
  }
}

Office.onReady(() => {
  $("generateRetainer").addEventListener("click", () => {
    generate("retainer");
  });

  $("generateCard").addEventListener("click", () => {
    generate("card");
  });

  showStatus(
    "Open a fresh copy of the appropriate Word template, " +
    "enter the information, and click its Generate button."
  );
});
