let xlsBtn = document.querySelector("#xls-button");
let csvBtn = document.querySelector("#csv-button");

const fetchContacts = async () => {
  const url =
    "https://script.googleusercontent.com/macros/echo?user_content_key=E3CVKUEg5N4oykJrbnPtheLMd8Ft0ifQ1Ly2lvGmN7qWc3HPNH9HnCI68-x-KYCK63sKSGN4V-EwxoYXKBHUzlcDDav4dF9xm5_BxDlH2jW0nuo2oDemN9CCS2h10ox_1xSncGQajx_ryfhECjZEnBfaEiyD8ASCT_NWu9uJBYEhKdEN-I354QK8XcksEa_YW2xfiQ5ILg04KPcR0q27pRS7xiu4BGT2VeodIqIma30teHllEL3Aptz9Jw9Md8uu&lib=M0RJuE2VVxuWLWyO6KVkgzVNpBywcptcy";

  const resp = await fetch(url);
  return resp.json()
};

const fetchCustomers = async () => {
  const url =
    "https://script.googleusercontent.com/macros/echo?user_content_key=0QXssGADwB7Dfieyag7cp3ojP5wqXxePoXDF3NwCJlbjXC0bodE-CUmdm7DLa5zxPDGuXZnHgbdirz_6EXryEvkA9ki7h_1qm5_BxDlH2jW0nuo2oDemN9CCS2h10ox_1xSncGQajx_ryfhECjZEnCALfLskPi0ipmYPrhHoIvcz_9_66QiYGrxNs0aeDMmW4Rf0WOm_un1RRfMjRHo__jDvyaFO0I7cvnOCf7aZzRvG2owbghyPpg&lib=MYaOfHpXjh5aDaEu1ffsz2hKSiheyHQNG";

  const resp = await fetch(url);
  return resp.json()
};

function generateOutputData(customers, contacts) {
  customers.shift();
  var newArray1 = [];

  for (var element of customers) {
    var similarityarray = [];

    for (var elements of contacts) {
      var a = elements["Name"];
      var similarity = stringSimilarity.compareTwoStrings(element, a);
      similarityarray.push(similarity);

      if (similarity >= 0.6) {
        var gbese = elements;
        newArray1.push(gbese);
      }

      if (similarityarray.length == contacts.length) {
        var result = similarityarray.every(function (e) {
          return e < 0.6;
        });

        if (result == true) {
          var comot = element;
          var outcasts = {};
          outcasts["Name"] = comot;
          outcasts["Phone"] = " ";
          newArray1.push(outcasts);
        }
      }
    }
  }

  return newArray1;
}

const start = async () => {
  console.log("Loading contacts...");
  const contacts = await fetchContacts();

  console.log("Loading customers...");
  const customers = (await fetchCustomers()).flat();
  const output = generateOutputData(customers, contacts);

  console.log('Done downloading data')
  var xls = new XlsExport(output, "data");
  console.log("XLS generated!")

  xlsBtn.addEventListener("click", function () {
    xls.exportToXLS("export-example.xls");
  });

  csvBtn.addEventListener("click", function () {
    xls.exportToCSV();
  });
};

start();
