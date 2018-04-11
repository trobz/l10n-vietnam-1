.. image:: https://img.shields.io/badge/licence-AGPL--3-blue.svg
   :target: http://www.gnu.org/licenses/agpl-3.0-standalone.html
   :alt: License: AGPL-3

Post Function One Time
============================
This application supports the system to execute specific functions
only one time after the first upgrade. The executed functions will not be
recalled during next upgrades unless you remove them from the 
"List_post_object_one_time_functions" system parameter.
There is one main function:
* run_post_object_one_time(object_name, list_functions=[])

Installation
============

Just simply install the module.

Usage
=====

Declare functions that you want to run during upgrade in XML file. 

Example:

.. code-block:: xml

	<openerp>
		<data noupdate="0">
			<function name="start" model="post.object.test.data" />
		</data>
	</openerp>


.. code-block:: python

	from openerp import api, models

	class PostObjectTestData(models.AbstractModel):
		_name = "post.object.test.data"
		_description = "Set Up Data"

		@api.model
		def start(self):
			# Functions run all times
			self.function_run_all_time_1()
			self.function_run_all_time_2()
			self.function_run_all_time_3()

			# Functions run one time
			PostFunction = self.env['post.function.one.time']
			PostFunction.run_post_object_one_time(
				'post.object.test.data',
				[
					'function_run_one_time_1',
					'function_run_one_time_2',
					'function_run_one_time_3'
				]
			)

			return True

Known issues / Roadmap
======================

* ...

Bug Tracker
===========

Bugs are tracked on `GitHub Issues
<https://github.com/OCA/{project_repo}/issues>`_. In case of trouble, please
check there if your issue has already been reported. If you spotted it first,
help us smashing it by providing a detailed and welcomed `feedback
<https://github.com/OCA/
{project_repo}/issues/new?body=module:%20
{module_name}%0Aversion:%20
{branch}%0A%0A**Steps%20to%20reproduce**%0A-%20...%0A%0A**Current%20behavior**%0A%0A**Expected%20behavior**>`_.

Credits
=======

Images
------

* Odoo Community Association: `Icon <https://github.com/OCA/maintainer-tools/blob/master/template/module/static/description/icon.svg>`_.

Contributors
------------

* Tuan Nguyen Duc <tuannd@trobz.com>

Maintainer
----------

.. image:: https://odoo-community.org/logo.png
   :alt: Odoo Community Association
   :target: https://odoo-community.org

This module is maintained by the OCA.

OCA, or the Odoo Community Association, is a nonprofit organization whose
mission is to support the collaborative development of Odoo features and
promote its widespread use.

To contribute to this module, please visit https://odoo-community.org.
