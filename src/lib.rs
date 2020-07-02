#![feature(ptr_internals, allocator_api, core_intrinsics, dropck_eyepatch, specialization)]
#![feature(internal_uninit_const, maybe_uninit_extra, maybe_uninit_slice)]

extern crate zip;

extern crate yaserde;

#[macro_use]
extern crate yaserde_derive;

mod xlsx;