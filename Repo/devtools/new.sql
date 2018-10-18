DECLARE @String nvarchar(400) = 'This is an IP Address: 10.0.80.1 That was an IP Address'

SELECT LEFT(
	SUBSTRING(
		@String, 
		PATINDEX(
			'%[0-9]%.%[0-9]%.%[0-9]%.%[0-9]%', 
			@String
		), 
		9999
	),
	PATINDEX(
		'%[^0-9]%', 
		SUBSTRING(
			@String, 
			PATINDEX(
				'%[0-9]%.%[0-9]%.%[0-9]%.%[0-9]%', 
				@String
			), 
			9999
		)
	)
)