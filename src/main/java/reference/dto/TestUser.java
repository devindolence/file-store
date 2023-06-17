package reference.dto;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import reference.file.ExcelDto;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class TestUser implements ExcelDto {
    private String name;
    private String age;

}
