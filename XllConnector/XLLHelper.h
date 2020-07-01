# pragma once

#include <cassert>
#include <comdef.h>
#include <comutil.h>
#include <vector>

namespace Helper {

	// Helper class to access a two-dimensional SAFEARRAY with element type
	// VARIANT. The wrapper always creates such an array when the UDF expects
	// an argument of type SAFEARRAY*.
	class VariantMatrixAccessor {
		SAFEARRAY *m_psa;
		VARIANT *m_data;
	public:
		explicit VariantMatrixAccessor(SAFEARRAY *psa) {
			assert(psa != nullptr);
			assert(SafeArrayGetDim(psa) == 2);

			VARTYPE vt;
			HRESULT hr = SafeArrayGetVartype(psa, &vt);
			if (FAILED(hr))
				throw _com_error(hr);
			assert(vt == VT_VARIANT);

			VARIANT *data;
			hr = SafeArrayAccessData(psa, (void**)&data);
			if (FAILED(hr))
				throw std::invalid_argument("Cannot access data");

			m_psa = psa;
			m_data = data;
		}

		~VariantMatrixAccessor() {
			if (m_psa != nullptr) {
				SafeArrayUnaccessData(m_psa);
				m_data = nullptr;
				m_psa = nullptr;
			}
		}

		size_t rows() const { return m_psa->rgsabound[0].cElements; }

		size_t columns() const { return m_psa->rgsabound[1].cElements; }

		size_t size() const { return rows()*columns(); }

		VARIANT& operator[](size_t index) {
			return m_data[index];
		}

		// zero-based
		VARIANT& operator()(size_t row, size_t column) {
			return m_data[row*columns() + column];
		}
	};

	

	std::vector<double> vectorFromSafeArray1Dim(SAFEARRAY *safeArray);

}  // namespace Helper
