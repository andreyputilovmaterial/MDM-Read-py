"""Tests for the pure row-name builders in ``row_name_maker``.

These run without the MDM COM runtime (no Windows / IBM Professional required).
"""

import sys
import unittest
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'src'))

import row_name_maker as rnm


class TestConventions(unittest.TestCase):
    def test_registry_has_ap_and_mdm(self):
        self.assertIn('ap', rnm.ROW_NAME_CONVENTIONS)
        self.assertIn('mdm', rnm.ROW_NAME_CONVENTIONS)

    def test_resolve_convention_returns_dict(self):
        self.assertEqual(rnm.resolve_convention('ap')['mode'], 'compose')
        self.assertEqual(rnm.resolve_convention('mdm')['mode'], 'fullname')

    def test_resolve_convention_unknown_raises(self):
        with self.assertRaises(ValueError):
            rnm.resolve_convention('nope')

    def test_resolve_convention_accepts_dict(self):
        d = {'mode': 'compose', 'separator': '.'}
        self.assertIs(rnm.resolve_convention(d), d)


class TestApConvention(unittest.TestCase):
    def test_root_name(self):
        item = SimpleNamespace(Name='Gender')
        self.assertEqual(rnm.make_root_name(item, 'ap'), 'Gender')

    def test_child_name(self):
        parent = SimpleNamespace(Name='Block')
        child = SimpleNamespace(Name='SubField')
        self.assertEqual(rnm.make_child_name(parent, child, 'ap'), 'Block.SubField')

    def test_child_name_accepts_parent_string(self):
        child = SimpleNamespace(Name='SubField')
        self.assertEqual(rnm.make_child_name('Block', child, 'ap'), 'Block.SubField')

    def test_category_name(self):
        parent = SimpleNamespace(Name='Q1')
        child = SimpleNamespace(Name='Male')
        self.assertEqual(rnm.make_category_name(parent, child, convention='ap'), 'Q1.Categories[Male]')

    def test_category_name_stepover_reference(self):
        parent = SimpleNamespace(Name='USLocationData.State')
        # In stepover mode the child_name is the referenced shared list's name.
        self.assertEqual(
            rnm.make_category_name(parent, SimpleNamespace(), child_name='SL_StateList', is_reference=True, convention='ap'),
            'USLocationData.State.Categories[@SL_StateList]',
        )

    def test_category_name_custom_reference_prefix(self):
        parent = SimpleNamespace(Name='Q1')
        self.assertEqual(
            rnm.make_category_name(parent, SimpleNamespace(), child_name='SL_X', is_reference=True, reference_prefix='#', convention='ap'),
            'Q1.Categories[#SL_X]',
        )

    def test_is_reference(self):
        self.assertTrue(rnm.is_reference(SimpleNamespace(IsReference=True)))
        self.assertFalse(rnm.is_reference(SimpleNamespace(IsReference=False)))
        self.assertFalse(rnm.is_reference(SimpleNamespace()))


class TestMdmConvention(unittest.TestCase):
    def test_root_name_uses_full_name(self):
        item = SimpleNamespace(Name='SubField', FullName='Block.SubField')
        self.assertEqual(rnm.make_root_name(item, 'mdm'), 'Block.SubField')

    def test_child_name_uses_full_name(self):
        child = SimpleNamespace(Name='SubField', FullName='Block.SubField')
        self.assertEqual(rnm.make_child_name('Block', child, 'mdm'), 'Block.SubField')

    def test_category_name_uses_full_name(self):
        child = SimpleNamespace(Name='Cat', FullName='Q1.Categories[Cat]')
        self.assertEqual(rnm.make_category_name('Q1', child, convention='mdm'), 'Q1.Categories[Cat]')

    def test_missing_full_name_raises_detailed_error(self):
        child = SimpleNamespace(Name='SubField')  # no FullName
        with self.assertRaises(AttributeError) as ctx:
            rnm.make_child_name('Block', child, convention='mdm')
        err = ctx.exception
        self.assertIn('Block.SubField', str(err))
        self.assertIn('FullName', str(err))
        # the original AttributeError must be chained via "from"
        self.assertIsNotNone(err.__cause__)
        self.assertIsInstance(err.__cause__, AttributeError)

    def test_missing_full_name_and_name_uses_raw_fallback(self):
        child = object()  # has neither Name nor FullName
        with self.assertRaises(AttributeError) as ctx:
            rnm.make_root_name(child, convention='mdm')
        err = ctx.exception
        self.assertIn('FullName', str(err))
        self.assertIsNotNone(err.__cause__)


if __name__ == '__main__':
    unittest.main()
