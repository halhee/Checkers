use pyo3::prelude::*;
use pyo3::types::{PyDict, PyList};
use rayon::prelude::*;
use std::collections::HashMap;
use serde::{Serialize, Deserialize};

#[derive(Serialize, Deserialize)]
struct ElementStats {
    total: usize,
    valid: usize,
    missing_pset: usize,
    missing_param: usize,
}

#[derive(Serialize, Deserialize)]
struct FloorStats {
    valid: usize,
    invalid: usize,
}

#[pyfunction]
fn analyze_elements(
    py: Python,
    elements: &PyList,
    required_psets: &PyDict,
    element_psets: &PyDict,
) -> PyResult<PyObject> {
    // Conversion des données Python en structures Rust
    let elements: Vec<String> = elements
        .iter()
        .map(|e| e.extract::<String>())
        .collect::<Result<Vec<_>, _>>()?;

    let required_psets: HashMap<String, HashMap<String, Vec<String>>> = required_psets
        .extract()?;

    let element_psets: HashMap<String, HashMap<String, HashMap<String, String>>> = element_psets
        .extract()?;

    // Statistiques globales
    let mut stats = ElementStats {
        total: elements.len(),
        valid: 0,
        missing_pset: 0,
        missing_param: 0,
    };

    // Statistiques par étage
    let mut floor_stats: HashMap<String, FloorStats> = HashMap::new();

    // Analyse parallèle des éléments
    let results: Vec<(bool, String, usize, usize)> = elements.par_iter()
        .map(|element_id| {
            let mut missing_pset = 0;
            let mut missing_param = 0;
            let mut is_valid = true;

            if let Some(element_data) = element_psets.get(element_id) {
                for (pset_name, required_params) in &required_psets {
                    if let Some(pset) = element_data.get(pset_name) {
                        for param_name in required_params.values().flatten() {
                            if !pset.contains_key(param_name) {
                                missing_param += 1;
                                is_valid = false;
                            }
                        }
                    } else {
                        missing_pset += 1;
                        is_valid = false;
                    }
                }
            }

            let floor = element_data
                .get("floor")
                .map(|f| f.to_string())
                .unwrap_or_else(|| "unknown".to_string());

            (is_valid, floor, missing_pset, missing_param)
        })
        .collect();

    // Agrégation des résultats
    for (is_valid, floor, mp, mm) in results {
        if is_valid {
            stats.valid += 1;
        }
        stats.missing_pset += mp;
        stats.missing_param += mm;

        let floor_stat = floor_stats
            .entry(floor)
            .or_insert(FloorStats { valid: 0, invalid: 0 });

        if is_valid {
            floor_stat.valid += 1;
        } else {
            floor_stat.invalid += 1;
        }
    }

    // Conversion des résultats en dictionnaire Python
    let result = PyDict::new(py);
    result.set_item("total_elements", stats.total)?;
    result.set_item("valid_elements", stats.valid)?;
    result.set_item("missing_psets", stats.missing_pset)?;
    result.set_item("missing_params", stats.missing_param)?;

    let floors: Vec<(&String, &FloorStats)> = floor_stats.iter().collect();
    let py_floors = PyList::empty(py);
    for (floor_name, stats) in floors {
        let floor_dict = PyDict::new(py);
        floor_dict.set_item("name", floor_name)?;
        floor_dict.set_item("valid", stats.valid)?;
        floor_dict.set_item("invalid", stats.invalid)?;
        py_floors.append(floor_dict.into())?;
    }
    result.set_item("floors", py_floors)?;

    Ok(result.into())
}

#[pymodule]
fn ifc_analyzer(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(analyze_elements, m)?)?;
    Ok(())
}
