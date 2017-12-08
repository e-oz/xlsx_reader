use zip;
use std::io::Cursor;
use std::io::Read;
use std::collections::HashMap;
use serde_xml::from_str;
use serde_xml::value::{Element, Content};
use std::char;

pub fn parse_xlsx(data: &Vec<u8>) -> Result<HashMap<usize, HashMap<usize, String>>, String> {
  let (strings, sheet) = match parse_xlsx_file_to_parts(data) {
    Ok(r) => r,
    Err(err) => return Err(err)
  };
  let map = match get_strings_map(strings) {
    Some(m) => m,
    None => return Err("Data extracting error".to_owned())
  };
  get_parsed_xlsx(map, sheet)
}

pub fn parse_xlsx_file_to_parts(data: &Vec<u8>) -> Result<(String, String), String>
{
  let reader = Cursor::new(data);
  let mut zip = match zip::ZipArchive::new(reader) {
    Ok(z) => z,
    Err(err) => return Err(format!("{:?}", err))
  };

  let mut strings_content = String::new();
  let mut sheet_content = String::new();

  for i in 0..zip.len() {
    let mut file = match zip.by_index(i) { Ok(f) => f, Err(_) => continue };
    if file.name() == "xl/sharedStrings.xml" {
      match file.read_to_string(&mut strings_content) {
        Ok(_) => (), Err(err) => return Err(format!("Can't read strings file: {:?}", err))
      }
    } else {
      if file.name() == "xl/worksheets/sheet1.xml" {
        match file.read_to_string(&mut sheet_content) {
          Ok(_) => (), Err(err) => return Err(format!("Can't read sheet file: {:?}", err))
        }
      }
    }
  }
  Ok((strings_content, sheet_content))
}

pub fn get_strings_map(strings: String) -> Option<HashMap<usize, String>>
{
  let xml_value: Element = match from_str(&strings) {
    Ok(c) => c,
    Err(_) => return None
  };
  let mut map: HashMap<usize, String> = HashMap::new();
  let mut i = 0;
  match xml_value.members {
    Content::Members(sst) => {
      if sst.contains_key("si") {
        for si in &sst["si"] {
          if si.attributes.contains_key("t") {
            map.insert(i, si.attributes["t"][0].clone());
          } else {
            map.insert(i, "".to_owned());
          }
          i = i + 1;
        }
      }
    },
    _ => ()
  }
  Some(map)
}

pub fn get_parsed_xlsx(strings_map: HashMap<usize, String>, sheet_content: String) -> Result<HashMap<usize, HashMap<usize, String>>, String>
{
  let xml_value: Element = match from_str(&sheet_content) {
    Ok(c) => c,
    Err(err) => return Err(format!("XML parsing error: {:?}", err))
  };

  match xml_value.members {
    Content::Members(worksheet) => {
      if worksheet.contains_key("sheetData") {
        let sheet_data = &worksheet["sheetData"][0];
        match sheet_data.members {
          Content::Members(ref rows_el) => {
            if rows_el.contains_key("row") {
              let mut table: HashMap<usize, HashMap<usize, String>> = HashMap::with_capacity(rows_el["row"].len());
              let mut ir: usize = 0;
              for row in &rows_el["row"] {
                match row.members {
                  Content::Members(ref cells) => {
                    if cells.contains_key("c") {
                      let mut tr: HashMap<usize, String> = HashMap::with_capacity(cells.len());
                      let mut i: usize = 0;
                      let cells_count = cells["c"].len();
                      for cell in &cells["c"] {
                        if cell.attributes.contains_key("r") {
                          let pre_i = i;
                          while excel_str_cell(ir + 1, i) != cell.attributes["r"][0] {
                            i += 1;
                            if i > cells_count {
                              i = pre_i;
                              break;
                            }
                          }
                        }
                        match cell.members {
                          Content::Text(ref t) => { tr.insert(i, t.clone()); },
                          _ => {
                            if cell.attributes.contains_key("v") {
                              if cell.attributes.contains_key("t") && cell.attributes["t"][0] == "s" {
                                let val = match cell.attributes["v"][0].parse::<usize>() {
                                  Ok(map_index) => {
                                    if strings_map.contains_key(&map_index) {
                                      strings_map[&map_index].clone()
                                    } else {
                                      "".to_owned()
                                    }
                                  },
                                  Err(_) => cell.attributes["v"][0].clone()
                                };
                                tr.insert(i, val);
                              } else {
                                if cell.attributes.contains_key("s") && (cell.attributes["s"][0] == "10" || cell.attributes["s"][0] == "14" || cell.attributes["s"][0] == "15") {
                                  tr.insert(i, excel_date(&cell.attributes["v"][0], Some(1462.0)));
                                } else {
                                  if cell.attributes.contains_key("s") && (cell.attributes["s"][0] == "4" || cell.attributes["s"][0] == "3" || cell.attributes["s"][0] == "5") {
                                    tr.insert(i, excel_date(&cell.attributes["v"][0], None));
                                  } else {
                                    tr.insert(i, cell.attributes["v"][0].clone());
                                  }
                                }
                              }
                            }
                          }
                        }
                        i = i + 1;
                      }
                      table.insert(ir, tr);
                    }
                  },
                  _ => ()
                }
                ir = ir + 1;
              }
              return Ok(table);
            }
          },
          _ => ()
        }
      }
    },
    _ => ()
  }
  Err("not impl".to_owned())
}

pub fn excel_date(src: &str, days_offset: Option<f64>) -> String {
  let mut days: f64 = match src.parse::<f64>() {
    Ok(i) => i + days_offset.unwrap_or(0.0),
    Err(_) => return "".to_owned()
  };
  let d: isize;
  let m: isize;
  let y: isize;
  if days == 60.0 {
    d = 29;
    m = 2;
    y = 1900;
  } else {
    if days < 60.0 {
      // Because of the 29-02-1900 bug, any serial date 
      // under 60 is one off... Compensate.
      days += 1.0;
    }
    // Modified Julian to DMY calculation with an addition of 2415019
    let mut l = (days as isize) + 68569 + 2415019;
    let n = (4 * l) / 146097;
    l = l - ((146097 * n + 3) / 4);
    let i = (4000 * (l + 1)) / 1461001;
    l = l - ((1461 * i) / 4) + 31;
    let j = (80 * l) / 2447;
    d = l - ((2447 * j) / 80);
    l = j / 11;
    m = j + 2 - (12 * l);
    y = 100 * (n - 49) + i + l;
  }
  return format!("{}-{:02}-{:02}", y, m, d)
}

pub fn excel_str_cell(row: usize, cell: usize) -> String {
  if cell == 0 {
    return format!("A{}", row);
  }
  let mut dividend = cell + 1;
  let mut column_name = String::new();
  let mut modulo;

  while dividend > 0 {
    modulo = (dividend - 1) % 26;
    column_name = format!("{}{}", char::from_u32((65 + modulo) as u32).unwrap_or('A'), column_name);
    dividend = (dividend - modulo) / 26;
  }

  format!("{}{}", column_name, row)
}
