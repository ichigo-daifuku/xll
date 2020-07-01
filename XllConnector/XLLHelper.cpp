#include "XLLHelper.h"

#include <typeinfo>

namespace Helper {

	template <typename T> T variant_cast(const VARIANT &);

	template <> double variant_cast<double>(const VARIANT &v) {
		VARIANT result;
		VariantInit(&result);
		HRESULT hr = VariantChangeType(&result, &v, 0, VT_R8);
		if (FAILED(hr))
			throw std::bad_cast();
		return V_R8(&result);
	}

	std::vector<double> vectorFromSafeArray1Dim(SAFEARRAY *safeArray) {
		VariantMatrixAccessor safeArrayAccessor(safeArray);
		size_t nSize = safeArrayAccessor.size();
		std::vector<double> result(nSize);
		for (size_t i = 0; i < (size_t)nSize; i++)
			result[i] = variant_cast<double>(safeArrayAccessor[i]);
		return result;
	}

}  // namespace Helper
