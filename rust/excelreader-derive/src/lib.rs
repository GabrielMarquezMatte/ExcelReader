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
    expand(input).unwrap_or_else(|e| e.to_compile_error()).into()
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

fn named_fields(input: &DeriveInput) -> syn::Result<&syn::punctuated::Punctuated<Field, syn::Token![,]>> {
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
    if let Type::Path(type_path) = ty {
        if let Some(segment) = type_path.path.segments.last() {
            if segment.ident == "Option" {
                if let PathArguments::AngleBracketed(args) = &segment.arguments {
                    if let Some(GenericArgument::Type(inner)) = args.args.first() {
                        return (inner, true);
                    }
                }
            }
        }
    }
    (ty, false)
}

enum FieldKind {
    Str,
    I64,
    F64,
    Bool,
}

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
        match name.as_deref() {
            Some("String") => Ok(FieldKind::Str),
            Some("i64") => Ok(FieldKind::I64),
            Some("f64") => Ok(FieldKind::F64),
            Some("bool") => Ok(FieldKind::Bool),
            _ => Err(syn::Error::new_spanned(
                ty,
                "unsupported field type for #[derive(ExcelMapper)] - supported types are \
                 String, i64, f64, bool, and Option<...> of those",
            )),
        }
    }

    fn xl_type_tokens(&self) -> proc_macro2::TokenStream {
        match self {
            FieldKind::Str => quote! { ::excelreader::XL_T_STRING },
            FieldKind::I64 => quote! { ::excelreader::XL_T_I64 },
            FieldKind::F64 => quote! { ::excelreader::XL_T_F64 },
            FieldKind::Bool => quote! { ::excelreader::XL_T_BOOL },
        }
    }

    fn value_tokens(&self) -> proc_macro2::TokenStream {
        match self {
            FieldKind::Str => quote! { ::excelreader::workbook::column_str(col, row).to_string() },
            FieldKind::I64 => quote! { ::excelreader::workbook::column_i64(col, row) },
            FieldKind::F64 => quote! { ::excelreader::workbook::column_f64(col, row) },
            FieldKind::Bool => quote! { ::excelreader::workbook::column_bool(col, row) },
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
                nome: u32,
            }
            "#,
        )
        .expect_err("must fail: u32 is not a supported type");

        assert!(err.to_string().contains("unsupported field type"));
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
