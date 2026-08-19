//! `#[derive(ExcelMapper)]` for excelreader's `ExcelMapper` trait. Do not depend on this crate
//! directly - `excelreader` re-exports the macro from `excelreader::workbook::ExcelMapper`
//! (same identifier as the trait; derive macros and traits occupy different namespaces, so
//! there's no collision, same as `serde`/`serde_derive`).

use proc_macro::TokenStream;
use quote::quote;
use syn::spanned::Spanned;
use syn::{
    parse_macro_input, Data, DeriveInput, Field, Fields, GenericArgument, LitStr, PathArguments,
    Type,
};

#[proc_macro_derive(ExcelMapper, attributes(excel))]
pub fn derive_excel_mapper(input: TokenStream) -> TokenStream {
    let input = parse_macro_input!(input as DeriveInput);
    expand(input)
        .unwrap_or_else(|e| e.to_compile_error())
        .into()
}

fn expand(input: DeriveInput) -> syn::Result<proc_macro2::TokenStream> {
    let struct_name = &input.ident;
    let fields = named_fields(&input)?;

    let bindings = fields
        .iter()
        .map(field_binding)
        .collect::<syn::Result<Vec<_>>>()?;

    Ok(quote! {
        impl ::excelreader::workbook::ExcelMapper for #struct_name {
            fn bindings() -> ::std::vec::Vec<::excelreader::workbook::ColumnBinding<Self>> {
                ::std::vec![ #(#bindings),* ]
            }
        }
    })
}

fn named_fields(
    input: &DeriveInput,
) -> syn::Result<&syn::punctuated::Punctuated<Field, syn::Token![,]>> {
    match &input.data {
        Data::Struct(data) => match &data.fields {
            Fields::Named(fields) => Ok(&fields.named),
            _ => Err(syn::Error::new_spanned(
                input,
                "ExcelMapper can only be derived for structs with named fields",
            )),
        },
        _ => Err(syn::Error::new_spanned(
            input,
            "ExcelMapper can only be derived for structs with named fields",
        )),
    }
}

fn field_binding(field: &Field) -> syn::Result<proc_macro2::TokenStream> {
    let field_ident = field.ident.as_ref().expect("named_fields guarantees Some");
    let name = excel_name(field)?;
    let (inner_ty, is_option) = unwrap_option(&field.ty);
    let kind = FieldKind::from_type(inner_ty)?;
    let xl_type = kind.xl_type_tokens();
    let value = kind.value_tokens();
    let assign_value = if is_option {
        quote! { ::std::option::Option::Some(#value) }
    } else {
        value
    };

    Ok(quote! {
        ::excelreader::workbook::ColumnBinding {
            name: #name,
            xl_type: #xl_type,
            assign: |r, col, row| r.#field_ident = #assign_value,
        }
    })
}

fn excel_name(field: &Field) -> syn::Result<LitStr> {
    for attr in &field.attrs {
        if !attr.path().is_ident("excel") {
            continue;
        }
        let mut name = None;
        attr.parse_nested_meta(|meta| {
            if meta.path.is_ident("name") {
                let value = meta.value()?;
                name = Some(value.parse::<LitStr>()?);
                Ok(())
            } else {
                Err(meta.error("unsupported #[excel(...)] key, expected `name`"))
            }
        })?;
        if let Some(name) = name {
            return Ok(name);
        }
    }
    Err(syn::Error::new(
        field.span(),
        "field is missing #[excel(name = \"...\")]",
    ))
}

/// Returns `(T, true)` for a field typed `Option<T>`, `(ty, false)` for anything else.
fn unwrap_option(ty: &Type) -> (&Type, bool) {
    let Type::Path(type_path) = ty else {
        return (ty, false);
    };
    let Some(segment) = type_path.path.segments.last() else {
        return (ty, false);
    };
    if segment.ident != "Option" {
        return (ty, false);
    }
    let PathArguments::AngleBracketed(args) = &segment.arguments else {
        return (ty, false);
    };
    if let Some(GenericArgument::Type(inner)) = args.args.first() {
        return (inner, true);
    }
    (ty, false)
}

enum FieldKind {
    Str,
    /// `i64` itself, and every other integer width. The column is always `XL_T_I64` on the wire; a
    /// narrower field converts through `TryFrom` (see `value_tokens`).
    Int,
    /// `f64` itself, plus `f32`. The column is always `XL_T_F64` on the wire.
    Float,
    Bool,
    Date,
    Time,
    Timestamp,
}

/// Every integer type that maps onto an `XL_T_I64` column. `i64` is included so the conversion in
/// `value_tokens` stays uniform - `TryFrom<i64> for i64` exists via the blanket `From` impl and
/// compiles away to nothing.
const INT_TYPES: &[&str] = &[
    "i8", "i16", "i32", "i64", "isize", "u8", "u16", "u32", "u64", "usize",
];

impl FieldKind {
    fn from_type(ty: &Type) -> syn::Result<FieldKind> {
        let name = match ty {
            Type::Path(type_path) => type_path
                .path
                .segments
                .last()
                .map(|segment| segment.ident.to_string()),
            _ => None,
        };
        let name = name.as_deref();
        if let Some(name) = name {
            if INT_TYPES.contains(&name) {
                return Ok(FieldKind::Int);
            }
        }
        match name {
            Some("String") => Ok(FieldKind::Str),
            Some("f32" | "f64") => Ok(FieldKind::Float),
            Some("bool") => Ok(FieldKind::Bool),
            // The `Naive*` names are chrono's, reached through the `excelreader/chrono` feature's
            // `From` impls. Matching on the type's last path segment means `chrono::NaiveDate`, a
            // `use`d `NaiveDate`, and an alias to either all resolve the same way - and if the
            // feature is off, the generated `.into()` fails to compile with the missing `From` impl
            // named, which points straight at the fix.
            Some("Date" | "NaiveDate") => Ok(FieldKind::Date),
            Some("Time" | "NaiveTime") => Ok(FieldKind::Time),
            Some("Timestamp" | "NaiveDateTime") => Ok(FieldKind::Timestamp),
            _ => Err(syn::Error::new_spanned(
                ty,
                "unsupported field type for #[derive(ExcelMapper)] - supported types are String, \
                 bool, any integer (i8..i64/isize, u8..u64/usize), f32, f64, the temporal newtypes \
                 Date/Time/Timestamp (or chrono's NaiveDate/NaiveTime/NaiveDateTime with the \
                 `chrono` feature), and Option<...> of any of those",
            )),
        }
    }

    fn xl_type_tokens(&self) -> proc_macro2::TokenStream {
        match self {
            FieldKind::Str => quote! { ::excelreader::XL_T_STRING },
            FieldKind::Int => quote! { ::excelreader::XL_T_I64 },
            FieldKind::Float => quote! { ::excelreader::XL_T_F64 },
            FieldKind::Bool => quote! { ::excelreader::XL_T_BOOL },
            FieldKind::Date => quote! { ::excelreader::XL_T_DATE },
            FieldKind::Time => quote! { ::excelreader::XL_T_TIME },
            FieldKind::Timestamp => quote! { ::excelreader::XL_T_TIMESTAMP },
        }
    }

    fn value_tokens(&self) -> proc_macro2::TokenStream {
        match self {
            FieldKind::Str => quote! { ::excelreader::workbook::column_str(col, row).to_string() },
            // Through `TryFrom`, not `as`: an `as` cast would silently wrap a value that does not
            // fit the declared field (a 70000 in a `u16` column becoming 4464), and a parser that
            // quietly changes the number it read is worse than one that stops.
            FieldKind::Int => quote! {
                ::core::convert::TryFrom::try_from(::excelreader::workbook::column_i64(col, row))
                    .expect("value does not fit this field's integer type")
            },
            // `as` is the only available narrowing for f64 -> f32, and it saturates to infinity
            // rather than wrapping, so there is no silent-corruption case to guard here.
            FieldKind::Float => quote! { ::excelreader::workbook::column_f64(col, row) as _ },
            FieldKind::Bool => quote! { ::excelreader::workbook::column_bool(col, row) },
            FieldKind::Date => {
                quote! { ::core::convert::Into::into(::excelreader::workbook::column_date(col, row)) }
            }
            FieldKind::Time => {
                quote! { ::core::convert::Into::into(::excelreader::workbook::column_time(col, row)) }
            }
            FieldKind::Timestamp => {
                quote! { ::core::convert::Into::into(::excelreader::workbook::column_timestamp(col, row)) }
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use super::expand;
    use syn::DeriveInput;

    fn expand_str(src: &str) -> syn::Result<String> {
        let input: DeriveInput = syn::parse_str(src).expect("test input must parse as an item");
        expand(input).map(|tokens| tokens.to_string())
    }

    #[test]
    fn generates_one_binding_per_field_with_inferred_types() {
        let output = expand_str(
            r#"
            struct Row {
                #[excel(name = "Nome")]
                nome: String,
                #[excel(name = "Idade")]
                idade: i64,
                #[excel(name = "Peso")]
                peso: f64,
                #[excel(name = "Ativo")]
                ativo: bool,
            }
            "#,
        )
        .expect("expand must succeed");

        assert!(output.contains("impl :: excelreader :: workbook :: ExcelMapper for Row"));
        assert!(output.contains("XL_T_STRING"));
        assert!(output.contains("XL_T_I64"));
        assert!(output.contains("XL_T_F64"));
        assert!(output.contains("XL_T_BOOL"));
        assert!(output.contains("\"Nome\""));
        assert!(output.contains("\"Idade\""));
        assert!(output.contains("\"Peso\""));
        assert!(output.contains("\"Ativo\""));
    }

    #[test]
    fn wraps_option_fields_in_some() {
        let output = expand_str(
            r#"
            struct Row {
                #[excel(name = "Ativo")]
                ativo: Option<bool>,
            }
            "#,
        )
        .expect("expand must succeed");

        assert!(output.contains("Option :: Some"));
        assert!(output.contains("XL_T_BOOL"));
    }

    #[test]
    fn errors_when_name_attribute_is_missing() {
        let err = expand_str(
            r#"
            struct Row {
                nome: String,
            }
            "#,
        )
        .expect_err("must fail: no #[excel(name = ...)]");

        assert!(err.to_string().contains("missing #[excel(name"));
    }

    #[test]
    fn errors_on_unsupported_field_type() {
        let err = expand_str(
            r#"
            struct Row {
                #[excel(name = "Nome")]
                nome: Vec<u8>,
            }
            "#,
        )
        .expect_err("must fail: Vec<u8> is not a supported type");

        assert!(err.to_string().contains("unsupported field type"));
    }

    #[test]
    fn maps_every_integer_width_onto_an_i64_column() {
        for ty in [
            "i8", "i16", "i32", "i64", "isize", "u8", "u16", "u32", "u64", "usize",
        ] {
            let output = expand_str(&format!(
                r#"struct Row {{ #[excel(name = "N")] n: {ty}, }}"#
            ))
            .unwrap_or_else(|e| panic!("{ty} must be supported: {e}"));

            assert!(output.contains("XL_T_I64"), "{ty} must use XL_T_I64");
            // Narrowing goes through TryFrom, never a silent `as` cast.
            assert!(output.contains("try_from"), "{ty} must convert via TryFrom");
        }
    }

    #[test]
    fn maps_both_float_widths_onto_an_f64_column() {
        for ty in ["f32", "f64"] {
            let output = expand_str(&format!(
                r#"struct Row {{ #[excel(name = "V")] v: {ty}, }}"#
            ))
            .unwrap_or_else(|e| panic!("{ty} must be supported: {e}"));
            assert!(output.contains("XL_T_F64"), "{ty} must use XL_T_F64");
        }
    }

    #[test]
    fn maps_temporal_types_onto_their_own_column_types() {
        for (ty, expected) in [
            ("Date", "XL_T_DATE"),
            ("NaiveDate", "XL_T_DATE"),
            ("Time", "XL_T_TIME"),
            ("NaiveTime", "XL_T_TIME"),
            ("Timestamp", "XL_T_TIMESTAMP"),
            ("NaiveDateTime", "XL_T_TIMESTAMP"),
        ] {
            let output = expand_str(&format!(
                r#"struct Row {{ #[excel(name = "D")] d: {ty}, }}"#
            ))
            .unwrap_or_else(|e| panic!("{ty} must be supported: {e}"));
            assert!(output.contains(expected), "{ty} must use {expected}");
        }
    }

    #[test]
    fn resolves_temporal_types_through_a_qualified_path() {
        let output = expand_str(
            r#"
            struct Row {
                #[excel(name = "D")]
                d: chrono::NaiveDate,
            }
            "#,
        )
        .expect("a qualified path must resolve by its last segment");

        assert!(output.contains("XL_T_DATE"));
    }

    #[test]
    fn errors_on_non_struct_input() {
        let err = expand_str(
            r#"
            enum Row {
                A,
            }
            "#,
        )
        .expect_err("must fail: enum is not a struct");

        assert!(err
            .to_string()
            .contains("can only be derived for structs with named fields"));
    }
}
